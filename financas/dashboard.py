"""Cálculos do dashboard trimestral.

Todos os valores são somados em módulo (Abs), então tanto faz se a planilha
lança despesas como positivas ou negativas: o grupo da conta define o sinal.
Contas marcadas "fora do DRE" (transferências entre contas) não entram.
"""
from collections import defaultdict
from decimal import Decimal

from django.db.models import Count, Sum
from django.db.models.functions import Abs

from .models import Lancamento
from .plano_contas import (COMISSOES, FORMAS_RECEBIMENTO, IMPOSTOS, LOGISTICA, MAO_DE_OBRA, MATERIA_PRIMA,
                           PROLABORE, grupo_do_codigo)

ZERO = Decimal(0)
MESES_ABREV = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]


def meses_do_periodo(ano, trimestre):
    if trimestre:
        return [(ano, m) for m in range(3 * trimestre - 2, 3 * trimestre + 1)]
    return [(ano, m) for m in range(1, 13)]


def periodo_anterior(ano, trimestre):
    if not trimestre:
        return ano - 1, None
    return (ano, trimestre - 1) if trimestre > 1 else (ano - 1, 4)


def rotulo(ano, trimestre):
    return f"T{trimestre}/{str(ano)[2:]}" if trimestre else str(ano)


def _pct(parte, todo):
    return (parte / todo * 100) if todo else None


def _n(v, casas=1):
    """Número no formato brasileiro: 1234.5 -> '1.234,5'."""
    return f"{v:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _var(atual, anterior):
    """Variação percentual; None quando não há base de comparação."""
    if atual is None or anterior in (None, ZERO):
        return None
    return (atual - anterior) / abs(anterior) * 100


class Base:
    """Totais por (ano, mês, código da conta), carregados uma vez."""

    def __init__(self, anos):
        self.totais = defaultdict(lambda: defaultdict(lambda: ZERO))  # (ano, mes) -> codigo -> valor
        self.nomes = {}
        self.ids = {}
        self.sem_conta = defaultdict(lambda: [0, ZERO])               # (ano, mes) -> [n, valor]
        self.sem_codigo = defaultdict(lambda: ZERO)                   # nome da conta -> valor
        self.sem_valor = defaultdict(int)
        self.n = defaultdict(int)
        linhas = (Lancamento.objects.filter(ano__in=anos)
                  .exclude(conta__nao_entra_no_dre=True)
                  .values("ano", "mes", "conta__codigo", "conta__nome", "conta_id")
                  .annotate(total=Sum(Abs("valor")), n=Count("id"), n_valor=Count("valor")))
        for l in linhas:
            chave = (l["ano"], l["mes"])
            total = l["total"] or ZERO
            self.n[chave] += l["n"]
            self.sem_valor[chave] += l["n"] - l["n_valor"]
            if l["conta__nome"] is None:
                self.sem_conta[chave][0] += l["n"]
                self.sem_conta[chave][1] += total
            elif grupo_do_codigo(l["conta__codigo"]) is None:
                self.sem_codigo[l["conta__nome"]] += total
            else:
                self.totais[chave][l["conta__codigo"]] += total
                self.nomes[l["conta__codigo"]] = l["conta__nome"]
                self.ids[l["conta__codigo"]] = l["conta_id"]

    def tem_dados(self, meses):
        return any(self.n.get(m) for m in meses)

    def soma(self, meses, codigos=None, grupo=None):
        total = ZERO
        for m in meses:
            for cod, v in self.totais.get(m, {}).items():
                if (codigos and cod in codigos) or (grupo and grupo_do_codigo(cod) == grupo):
                    total += v
        return total

    def por_conta(self, meses, grupos):
        r = defaultdict(lambda: ZERO)
        for m in meses:
            for cod, v in self.totais.get(m, {}).items():
                if grupo_do_codigo(cod) in grupos:
                    r[cod] += v
        return r

    def indicadores(self, meses):
        s = self.soma
        receita = s(meses, grupo=1)
        cv = s(meses, grupo=3)
        cf = s(meses, grupo=5)
        mc = receita - cv
        resultado = mc - cf
        rec_no, desp_no = s(meses, grupo=2), s(meses, grupo=7)
        mc_pct = _pct(mc, receita)
        ind = {
            "receita": receita, "cv": cv, "mc": mc, "mc_pct": mc_pct, "cf": cf,
            "resultado": resultado, "resultado_pct": _pct(resultado, receita),
            "rec_nao_oper": rec_no, "desp_nao_oper": desp_no,
            "saldo_final": resultado + rec_no - desp_no,
            "ponto_equilibrio": (cf / (mc_pct / 100)) if mc_pct and mc_pct > 0 else None,
            "logistica": s(meses, LOGISTICA), "comissoes": s(meses, COMISSOES),
            "impostos": s(meses, IMPOSTOS), "mao_obra": s(meses, MAO_DE_OBRA),
            "prolabore": s(meses, PROLABORE),
            "tem_dados": self.tem_dados(meses),
        }
        for cod, nome in MATERIA_PRIMA.items():
            ind["mp_" + cod] = s(meses, (cod,))
        for chave in ["logistica", "impostos", "mao_obra", "comissoes", "mp_3.1", "mp_3.2", "mp_3.7"]:
            ind[chave + "_pct"] = _pct(ind[chave], receita)
        ind["cobertura_pe"] = _pct(receita, ind["ponto_equilibrio"]) if ind["ponto_equilibrio"] else None
        return ind


# Cartões: (chave, rótulo, formato, "bom" quando sobe?)
CARTOES = [
    ("receita", "Receita de vendas", "R$", True),
    ("cv", "Custos variáveis", "R$", False),
    ("mc", "Margem de contribuição", "R$", True),
    ("mc_pct", "Margem de contribuição %", "%", True),
    ("cf", "Custos fixos", "R$", False),
    ("resultado", "Resultado operacional", "R$", True),
    ("resultado_pct", "Resultado % da receita", "%", True),
    ("ponto_equilibrio", "Ponto de equilíbrio", "R$", False),
]

INDICADORES_PCT = [
    ("mp_3.1_pct", "Aço"), ("mp_3.2_pct", "Cimento"), ("mp_3.7_pct", "Mercadorias diversas"),
    ("logistica_pct", "Logística (frete + combustível)"), ("mao_obra_pct", "Mão de obra"),
    ("impostos_pct", "Impostos"),
]


def _cartoes(atual, anterior, ano_ant):
    cartoes = []
    for chave, rot, fmt, bom_sobe in CARTOES:
        v = atual[chave]
        def comp(base):
            if not base["tem_dados"] or v is None or base[chave] is None:
                return None
            delta = (v - base[chave]) if fmt == "%" else _var(v, base[chave])
            if delta is None:
                return None
            return {"delta": delta, "unidade": "p.p." if fmt == "%" else "%",
                    "bom": (delta >= 0) == bom_sobe if delta else None}
        cartoes.append({"chave": chave, "rotulo": rot, "valor": v, "formato": fmt,
                        "vs_anterior": comp(anterior), "vs_ano_anterior": comp(ano_ant)})
    return cartoes


def _alertas(base, meses, meses_ant, atual, anterior, ano_ant, rot_ant):
    alertas = []  # (prioridade, nível, texto)
    if not atual["tem_dados"]:
        return []
    if atual["receita"] and atual["resultado"] < 0:
        alertas.append((0, "ruim", "Prejuízo operacional no período: os custos superaram a receita."))
    if anterior["tem_dados"]:
        if atual["mc_pct"] is not None and anterior["mc_pct"] is not None:
            d = atual["mc_pct"] - anterior["mc_pct"]
            if abs(d) >= 3:
                alertas.append((1, "ruim" if d < 0 else "bom",
                                f"Margem de contribuição {'caiu' if d < 0 else 'subiu'} {_n(abs(d))} p.p. em relação a {rot_ant}."))
        for chave, nome in INDICADORES_PCT:
            a, b = atual[chave], anterior[chave]
            if a is not None and b is not None and abs(a - b) >= 2:
                alertas.append((2, "ruim" if a > b else "bom",
                                f"{nome} passou de {_n(b)}% para {_n(a)}% da receita (vs {rot_ant})."))
        lim = (atual["receita"] or ZERO) * Decimal("0.01") or Decimal(1000)
        cont_at, cont_ant = base.por_conta(meses, {3, 5}), base.por_conta(meses_ant, {3, 5})
        for cod, v in sorted(cont_at.items(), key=lambda kv: -kv[1]):
            b = cont_ant.get(cod, ZERO)
            var = _var(v, b)
            if var is not None and var >= 30 and v - b >= lim:
                alertas.append((3, "ruim", f"{base.nomes[cod].lstrip('-> ')} subiu {_n(var, 0)}% em relação a {rot_ant}."))
    if ano_ant["tem_dados"] and ano_ant["receita"]:
        var = _var(atual["receita"], ano_ant["receita"])
        if var is not None and abs(var) >= 15:
            alertas.append((1, "ruim" if var < 0 else "bom",
                            f"Receita {'caiu' if var < 0 else 'cresceu'} {_n(abs(var), 0)}% em relação ao mesmo período do ano anterior."))
    alertas.sort(key=lambda a: a[0])
    return [{"nivel": n, "texto": t} for _, n, t in alertas[:8]]


def _qualidade(base, meses):
    n = sum(base.n.get(m, 0) for m in meses)
    sem_n = sum(base.sem_conta[m][0] for m in meses if m in base.sem_conta)
    sem_v = sum((base.sem_conta[m][1] for m in meses if m in base.sem_conta), ZERO)
    total_v = sum((v for m in meses for v in base.totais.get(m, {}).values()), ZERO) + sem_v
    return {
        "n": n, "sem_conta_n": sem_n, "sem_conta_valor": sem_v,
        "pct_classificado_valor": _pct(total_v - sem_v, total_v) if total_v else None,
        "sem_valor": sum(base.sem_valor.get(m, 0) for m in meses),
        "contas_sem_codigo": sorted(base.sem_codigo),
    }


def montar(ano, trimestre=None):
    base = Base([ano - 2, ano - 1, ano])
    meses = meses_do_periodo(ano, trimestre)
    a_ant, t_ant = periodo_anterior(ano, trimestre)
    meses_ant = meses_do_periodo(a_ant, t_ant)
    meses_aa = meses_do_periodo(ano - 1, trimestre)
    atual, anterior, ano_ant = (base.indicadores(m) for m in (meses, meses_ant, meses_aa))
    rot_ant = rotulo(a_ant, t_ant)

    # Tendência: até 8 trimestres terminando no trimestre escolhido (ou no T4 do ano).
    fim_a, fim_t = ano, trimestre or 4
    trimestres = []
    a, t = fim_a, fim_t
    for _ in range(8):
        trimestres.insert(0, (a, t))
        a, t = (a, t - 1) if t > 1 else (a - 1, 4)
    while trimestres and not base.tem_dados(meses_do_periodo(*trimestres[0])):
        trimestres.pop(0)
    while trimestres and not base.tem_dados(meses_do_periodo(*trimestres[-1])):
        trimestres.pop()
    ind_tri = [base.indicadores(meses_do_periodo(a, t)) for a, t in trimestres]

    def serie(chave, inds):
        # Períodos sem lançamentos ficam vazios no gráfico (não como zero).
        return [float(round(i[chave], 2)) if i[chave] is not None and i["tem_dados"] else None for i in inds]

    tendencia = {
        "rotulos": [rotulo(a, t) for a, t in trimestres],
        "receita": serie("receita", ind_tri), "cv": serie("cv", ind_tri), "cf": serie("cf", ind_tri),
        "resultado": serie("resultado", ind_tri),
        "mc_pct": serie("mc_pct", ind_tri), "resultado_pct": serie("resultado_pct", ind_tri),
        "materia_prima": {nome: serie("mp_" + cod, ind_tri) for cod, nome in MATERIA_PRIMA.items()},
        "pct_receita": {nome: serie(chave, ind_tri) for chave, nome in INDICADORES_PCT},
    }

    ind_mes = [base.indicadores([(ano, m)]) for m in range(1, 13)]
    ind_mes_aa = [base.indicadores([(ano - 1, m)]) for m in range(1, 13)]
    mensal = {
        "rotulos": MESES_ABREV,
        "receita": serie("receita", ind_mes), "resultado": serie("resultado", ind_mes),
        "receita_aa": serie("receita", ind_mes_aa),
        "destaque": [m for _, m in meses],
        "tem_aa": any(i["tem_dados"] for i in ind_mes_aa),
    }

    # Maiores despesas (custos variáveis e fixos) do período, com variação.
    cont_at, cont_ant = base.por_conta(meses, {3, 5}), base.por_conta(meses_ant, {3, 5})
    total_desp = (atual["cv"] + atual["cf"]) or None
    top = []
    for cod, v in sorted(cont_at.items(), key=lambda kv: -kv[1])[:10]:
        top.append({"codigo": cod, "conta_id": base.ids[cod], "nome": base.nomes[cod].lstrip("-> "), "valor": v,
                    "pct": _pct(v, total_desp) if total_desp else None,
                    "anterior": cont_ant.get(cod), "var": _var(v, cont_ant.get(cod)) if anterior["tem_dados"] else None})

    formas = []
    for cod, nome in FORMAS_RECEBIMENTO.items():
        v = base.soma(meses, (cod,))
        if v:
            formas.append({"nome": nome, "valor": float(v), "pct": _pct(v, atual["receita"])})

    graficos = {
        "tendencia": tendencia, "mensal": mensal,
        "top": {"rotulos": [t["nome"] for t in top], "valores": [float(t["valor"]) for t in top]},
        "formas": {"rotulos": [f["nome"] for f in formas], "valores": [f["valor"] for f in formas]},
    }

    return {
        "ano": ano, "trimestre": trimestre, "rotulo": rotulo(ano, trimestre), "rotulo_anterior": rot_ant,
        "rotulo_ano_anterior": rotulo(ano - 1, trimestre),
        "atual": atual, "cartoes": _cartoes(atual, anterior, ano_ant),
        "indicadores_pct": [{"nome": n, "atual": atual[c], "anterior": anterior[c] if anterior["tem_dados"] else None,
                             "ano_anterior": ano_ant[c] if ano_ant["tem_dados"] else None}
                            for c, n in INDICADORES_PCT],
        "top": top, "formas": formas,
        "alertas": _alertas(base, meses, meses_ant, atual, anterior, ano_ant, rot_ant),
        "qualidade": _qualidade(base, meses),
        "graficos": graficos,
    }
