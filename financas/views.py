from collections import defaultdict
from datetime import datetime, time
from decimal import Decimal

from django.contrib import messages
from django.contrib.auth import get_user_model
from django.core.paginator import Paginator
from django.db.models import Count, Q, Sum
from django.http import FileResponse, HttpResponseBadRequest
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse
from django.utils import timezone
from django.views.decorators.http import require_POST

from .classificador import classificar, normalizar, reclassificar, regra_serve, regras_ativas
from .forms import MESES_FILTRO, ContaForm, ImportarForm, LancamentoForm, RegraForm
from .models import MESES, Conta, Lancamento, Regra, Secao
from . import planilha


def _anos():
    anos = list(Lancamento.objects.values_list("ano", flat=True).distinct().order_by("-ano"))
    atual = timezone.localdate().year
    if atual not in anos:
        anos.insert(0, atual)
    return anos


def _ano_padrao():
    return Lancamento.objects.order_by("-ano").values_list("ano", flat=True).first() or timezone.localdate().year


# ---------------------------------------------------------------- lançamentos

def lancamentos(request):
    g = request.GET
    ano = int(g.get("ano") or _ano_padrao())
    qs = Lancamento.objects.filter(ano=ano).select_related("conta", "regra")
    if g.get("mes"):
        qs = qs.filter(mes=g["mes"])
    if g.get("secao"):
        qs = qs.filter(secao=g["secao"])
    if g.get("conta"):
        qs = qs.filter(conta_id=g["conta"])
    if g.get("classificacao"):
        qs = qs.filter(classificacao=g["classificacao"])
    if g.get("q"):
        qs = qs.filter(descricao__icontains=g["q"].strip())

    totais = qs.aggregate(
        n=Count("id"), sem=Count("id", filter=Q(classificacao="N")),
        manuais=Count("id", filter=Q(classificacao="M")), soma=Sum("valor"),
    )
    pagina = Paginator(qs, 100).get_page(g.get("page"))
    params = g.copy()
    params.pop("page", None)
    return render(request, "financas/lancamentos.html", {
        "pagina": pagina, "totais": totais, "ano": ano, "anos": _anos(),
        "meses": MESES_FILTRO, "secoes": Secao.choices,
        "classificacoes": Lancamento.Classificacao.choices,
        "contas": Conta.objects.filter(ativa=True), "f": g, "querystring": params.urlencode(),
    })


def lancamento_form(request, pk=None):
    lanc = get_object_or_404(Lancamento, pk=pk) if pk else None
    form = LancamentoForm(request.POST or None, instance=lanc)
    if request.method == "POST" and form.is_valid():
        obj = form.save()
        txt = f"conta: {obj.conta}" if obj.conta else "sem regra que sirva — escolha a conta ou cadastre uma regra"
        messages.success(request, f"Lançamento salvo ({txt}).")
        if "salvar_e_novo" in request.POST:
            return redirect(f"{reverse('lancamento_novo')}?ano={obj.ano}&mes={obj.mes}&secao={obj.secao}")
        return redirect(f"{reverse('lancamentos')}?ano={obj.ano}&mes={obj.mes}")
    if not pk and request.method == "GET":
        for campo in ("ano", "mes", "secao"):
            if request.GET.get(campo):
                form.initial[campo] = request.GET[campo]
    historico = lanc.history.select_related("history_user")[:30] if lanc else []
    return render(request, "financas/lancamento_form.html", {
        "form": form, "lanc": lanc, "historico": _com_diferencas(historico),
    })


@require_POST
def lancamento_excluir(request, pk):
    lanc = get_object_or_404(Lancamento, pk=pk)
    ano, mes = lanc.ano, lanc.mes
    lanc.delete()
    messages.success(request, "Lançamento excluído (continua no histórico).")
    return redirect(f"{reverse('lancamentos')}?ano={ano}&mes={mes}")


@require_POST
def lancamento_conta(request, pk):
    """HTMX: troca a conta direto na tabela. Valor 'auto' devolve para as regras."""
    lanc = get_object_or_404(Lancamento.objects.select_related("conta"), pk=pk)
    valor = request.POST.get("conta", "")
    if valor == "auto":
        lanc.classificacao = Lancamento.Classificacao.NENHUMA
        lanc.aplicar_regras()
        lanc._change_reason = "Voltou para classificação automática"
    elif valor:
        lanc.conta = get_object_or_404(Conta, pk=valor)
        lanc.regra = None
        lanc.classificacao = Lancamento.Classificacao.MANUAL
        lanc._change_reason = "Correção manual da conta"
    else:
        return HttpResponseBadRequest("conta obrigatória")
    lanc.save()
    lanc.refresh_from_db()
    return render(request, "financas/_linha.html", {"l": lanc, "contas": Conta.objects.filter(ativa=True)})


def sugerir(request):
    """HTMX: mostra qual conta as regras dariam para uma descrição."""
    desc = request.GET.get("descricao", "")
    secao = request.GET.get("secao", "D")
    regra = classificar(desc, secao) if desc.strip() else None
    return render(request, "financas/_sugestao.html", {"regra": regra, "desc": desc})


@require_POST
def reclassificar_tudo(request):
    n = reclassificar(motivo="Reclassificação manual (botão)", usuario=request.user)
    messages.success(request, f"Regras reaplicadas: {n} lançamento(s) mudaram de conta.")
    return redirect(request.POST.get("next") or "lancamentos")


# ---------------------------------------------------------------- regras

def regras(request):
    qs = Regra.objects.select_related("conta").annotate()
    q = request.GET.get("q", "").strip()
    if q:
        qs = qs.filter(Q(palavra__icontains=q) | Q(conta__nome__icontains=q) | Q(observacao__icontains=q))
    if request.GET.get("conta"):
        qs = qs.filter(conta_id=request.GET["conta"])
    uso = dict(Lancamento.objects.filter(classificacao="A").values_list("regra").annotate(n=Count("id")))
    lista = list(qs.order_by("-ordem", "-id"))
    for r in lista:
        r.uso = uso.get(r.id, 0)

    # Testador: qual regra pega esta descrição e quais outras também serviriam.
    teste = None
    if request.GET.get("testar"):
        d, s = request.GET["testar"], request.GET.get("secao_teste", "D")
        ativas = regras_ativas()
        dn = normalizar(d)
        todas = [r for r in ativas if regra_serve(r, dn, s)]
        teste = {"descricao": d, "secao": s, "vencedora": todas[0] if todas else None, "outras": todas[1:]}

    return render(request, "financas/regras.html", {
        "regras": lista, "q": q, "teste": teste, "contas": Conta.objects.filter(ativa=True),
        "conta_sel": request.GET.get("conta", ""),
        "sem_regra": Lancamento.objects.filter(classificacao="N").values("descricao", "secao")
        .annotate(n=Count("id")).order_by("-n")[:30],
    })


def regra_form(request, pk=None):
    regra = get_object_or_404(Regra, pk=pk) if pk else None
    inicial = {}
    if not pk and request.GET.get("palavra"):
        inicial = {"palavra": request.GET["palavra"].upper(), "secao": request.GET.get("secao", "A")}
    form = RegraForm(request.POST or None, instance=regra, initial=inicial)
    if request.method == "POST" and form.is_valid():
        r = form.save()
        n = reclassificar(motivo=f"Regra {'alterada' if pk else 'criada'}: {r.palavra}", usuario=request.user)
        messages.success(request, f'Regra "{r.palavra}" salva. {n} lançamento(s) reclassificado(s).')
        return redirect("regras")
    historico = regra.history.select_related("history_user")[:30] if regra else []
    afetados = None
    if regra:
        afetados = Lancamento.objects.filter(regra=regra).order_by("-ano", "-mes")[:20]
    return render(request, "financas/regra_form.html", {
        "form": form, "regra": regra, "historico": _com_diferencas(historico), "afetados": afetados,
    })


@require_POST
def regra_excluir(request, pk):
    regra = get_object_or_404(Regra, pk=pk)
    nome = regra.palavra
    regra.delete()
    n = reclassificar(motivo=f"Regra excluída: {nome}", usuario=request.user)
    messages.success(request, f'Regra "{nome}" excluída. {n} lançamento(s) reclassificado(s).')
    return redirect("regras")


# ---------------------------------------------------------------- contas

def contas(request):
    lista = Conta.objects.annotate(n_regras=Count("regras", distinct=True),
                                   n_lanc=Count("lancamentos", distinct=True))
    return render(request, "financas/contas.html", {"contas": lista})


def conta_form(request, pk=None):
    conta = get_object_or_404(Conta, pk=pk) if pk else None
    form = ContaForm(request.POST or None, instance=conta)
    if request.method == "POST" and form.is_valid():
        form.save()
        messages.success(request, "Conta salva.")
        return redirect("contas")
    return render(request, "financas/conta_form.html", {"form": form, "conta": conta})


# ---------------------------------------------------------------- resumo

def resumo(request):
    ano = int(request.GET.get("ano") or _ano_padrao())
    linhas = (Lancamento.objects.filter(ano=ano)
              .values("secao", "conta__codigo", "conta__nome", "conta__nao_entra_no_dre", "mes")
              .annotate(total=Sum("valor"), n=Count("id")))
    tabela = defaultdict(lambda: {"meses": [Decimal(0)] * 12, "n": 0})
    tem_valor = False
    for l in linhas:
        chave = (l["secao"], l["conta__codigo"] or "", l["conta__nome"] or "(sem conta)",
                 l["conta__nao_entra_no_dre"] or False)
        if l["total"] is not None:
            tem_valor = True
            tabela[chave]["meses"][l["mes"] - 1] += l["total"]
        tabela[chave]["n"] += l["n"]
    secoes = []
    for sec, rotulo in Secao.choices:
        itens = sorted(((k, v) for k, v in tabela.items() if k[0] == sec), key=lambda kv: (kv[0][1], kv[0][2]))
        tot = [sum(v["meses"][i] for k, v in itens if not k[3]) for i in range(12)]
        secoes.append({
            "rotulo": rotulo,
            "linhas": [{"codigo": k[1], "nome": k[2], "fora": k[3], "meses": v["meses"],
                        "total": sum(v["meses"]), "n": v["n"]} for k, v in itens],
            "total_meses": tot, "total": sum(tot),
        })
    return render(request, "financas/resumo.html", {
        "ano": ano, "anos": _anos(), "secoes": secoes, "meses": [n for _, n in MESES], "tem_valor": tem_valor,
    })


# ---------------------------------------------------------------- dashboard

def dashboard(request):
    from . import dashboard as dash

    ano = int(request.GET.get("ano") or _ano_padrao())
    tri = request.GET.get("tri", "")
    if tri not in ("1", "2", "3", "4", "ano"):
        # padrão: trimestre do último mês com lançamentos no ano
        ultimo = Lancamento.objects.filter(ano=ano).order_by("-mes").values_list("mes", flat=True).first()
        tri = str((ultimo - 1) // 3 + 1) if ultimo else "ano"
    trimestre = None if tri == "ano" else int(tri)
    dados = dash.montar(ano, trimestre)
    tem_valores = Lancamento.objects.filter(ano=ano, valor__isnull=False).exists()
    return render(request, "financas/dashboard.html", {
        **dados, "anos": _anos(), "tri": tri, "tem_valores": tem_valores,
        "opcoes_tri": [("1", "T1"), ("2", "T2"), ("3", "T3"), ("4", "T4"), ("ano", "Ano")],
    })


# ---------------------------------------------------------------- histórico

MODELOS = {"lancamento": Lancamento, "regra": Regra, "conta": Conta}
TIPO_ALTERACAO = {"+": "Criação", "~": "Alteração", "-": "Exclusão"}


def _nomes_fk():
    return {
        "conta": dict(Conta.objects.values_list("id", "nome")),
        "regra": dict(Regra.objects.values_list("id", "palavra")),
    }


def _com_diferencas(registros, nomes=None):
    """Acrescenta a cada registro do histórico a lista de campos alterados."""
    registros = list(registros)
    nomes = nomes or _nomes_fk()
    for r in registros:
        r.tipo_txt = TIPO_ALTERACAO.get(r.history_type, r.history_type)
        r.mudancas = []
        if r.history_type == "~":
            anterior = r.prev_record
            if anterior:
                for m in r.diff_against(anterior).changes:
                    campo = r.instance._meta.get_field(m.field)
                    def fmt(v, campo=campo, nome=m.field):
                        if v in (None, ""):
                            return "—"
                        if nome in nomes:
                            return nomes[nome].get(v, f"#{v} (excluída)")
                        if campo.choices:
                            return dict(campo.choices).get(v, v)
                        return v
                    r.mudancas.append((campo.verbose_name, fmt(m.old), fmt(m.new)))
    return registros


def historico(request):
    g = request.GET
    modelos = [g["modelo"]] if g.get("modelo") in MODELOS else list(MODELOS)
    usuarios = get_user_model().objects.order_by("username")
    registros = []
    for nome in modelos:
        qs = MODELOS[nome].history.select_related("history_user")
        if g.get("usuario"):
            qs = qs.filter(history_user_id=g["usuario"])
        if g.get("tipo") in TIPO_ALTERACAO:
            qs = qs.filter(history_type=g["tipo"])
        if g.get("de"):
            qs = qs.filter(history_date__gte=timezone.make_aware(datetime.combine(
                datetime.fromisoformat(g["de"]).date(), time.min)))
        if g.get("ate"):
            qs = qs.filter(history_date__lte=timezone.make_aware(datetime.combine(
                datetime.fromisoformat(g["ate"]).date(), time.max)))
        if g.get("q"):
            campo = {"lancamento": "descricao", "regra": "palavra", "conta": "nome"}[nome]
            qs = qs.filter(Q(**{f"{campo}__icontains": g["q"]}) | Q(history_change_reason__icontains=g["q"]))
        for r in qs.order_by("-history_date")[:3000]:
            r.modelo = nome
            registros.append(r)
    registros.sort(key=lambda r: r.history_date, reverse=True)
    pagina = Paginator(registros, 50).get_page(g.get("page"))
    pagina.object_list = _com_diferencas(pagina.object_list)
    params = g.copy()
    params.pop("page", None)
    return render(request, "financas/historico.html", {
        "pagina": pagina, "usuarios": usuarios, "f": g, "querystring": params.urlencode(),
        "tipos": TIPO_ALTERACAO,
    })


# ---------------------------------------------------------------- planilha

def importar(request):
    form = ImportarForm(request.POST or None, request.FILES or None)
    resultado = None
    if request.method == "POST" and form.is_valid():
        d = form.cleaned_data
        if not (d["importar_regras"] or d["importar_lancamentos"]):
            form.add_error(None, "Marque pelo menos uma opção para importar.")
        else:
            try:
                wb = planilha.abrir_planilha(d["arquivo"], d["senha"])
                resultado = planilha.importar(
                    wb, d["ano"], usuario=request.user, importar_regras=d["importar_regras"],
                    importar_lancamentos=d["importar_lancamentos"], coluna_valor=d["coluna_valor"],
                )
                messages.success(request, "Importação concluída.")
            except planilha.ErroImportacao as exc:
                form.add_error(None, str(exc))
    return render(request, "financas/importar.html", {"form": form, "resultado": resultado})


def exportar(request):
    ano = int(request.GET.get("ano") or _ano_padrao())
    return FileResponse(planilha.exportar(ano), as_attachment=True,
                        filename=f"Gestao_Financeira_{ano}_export.xlsx")
