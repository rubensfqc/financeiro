"""Importação e exportação da planilha Gestão Financeira (.xlsx).

Estrutura esperada (a mesma da planilha com classificação automática):
- Aba "Regras": cabeçalho na linha 4, regras a partir da linha 5
  (A Palavra-chave | B Tipo de busca | C Seção | D Plano de Contas | E Observação),
  e a lista de contas válidas na coluna H a partir da linha 5.
- Aba "2024R" (opcional): código da conta na coluna A, nome na coluna B.
- Abas de mês Jan…Dez: descrição na coluna B, conta na coluna C, dados a partir
  da linha 4. Receitas acima da linha "DESPESAS", despesas abaixo dela.
"""
import io
import re
from decimal import Decimal, InvalidOperation

import openpyxl
from django.db import transaction
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import column_index_from_string, get_column_letter
from simple_history.utils import bulk_create_with_history

from .classificador import regras_ativas
from .models import MESES, Conta, Lancamento, Regra

ABAS_MESES = [nome for _, nome in MESES]
TIPOS = {"CONTÉM": "C", "CONTEM": "C", "COMEÇA COM": "I", "COMECA COM": "I"}
SECOES = {"RECEITA": "R", "DESPESA": "D", "AMBAS": "A"}
CABECALHOS = {"DESPESAS", "RECEITAS", "DESCRIÇÃO", "DESCRICAO"}


class ErroImportacao(Exception):
    pass


def abrir_planilha(arquivo, senha=""):
    """Abre o .xlsx, removendo a senha de abertura se houver."""
    dados = arquivo.read() if hasattr(arquivo, "read") else open(arquivo, "rb").read()
    buf = io.BytesIO(dados)
    try:
        import msoffcrypto

        office = msoffcrypto.OfficeFile(buf)
        if office.is_encrypted():
            if not senha:
                raise ErroImportacao("A planilha tem senha. Informe a senha para importar.")
            saida = io.BytesIO()
            try:
                office.load_key(password=senha)
                office.decrypt(saida)
            except Exception as exc:
                raise ErroImportacao("Senha da planilha incorreta.") from exc
            buf = saida
    except ErroImportacao:
        raise
    except Exception:
        buf.seek(0)  # não é um arquivo criptografado; segue normal
    buf.seek(0)
    try:
        # data_only=True: lê o valor calculado das fórmulas (ex.: a coluna C).
        return openpyxl.load_workbook(buf, data_only=True)
    except Exception as exc:
        raise ErroImportacao(f"Não consegui abrir o arquivo como .xlsx ({exc}).") from exc


def _texto(v):
    return str(v).strip() if v is not None else ""


def _decimal(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float, Decimal)):
        return Decimal(str(v)).quantize(Decimal("0.01"))
    s = re.sub(r"[^\d,.-]", "", str(v))
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return Decimal(s).quantize(Decimal("0.01"))
    except InvalidOperation:
        return None


def _codigos_contas(wb):
    """nome da conta → código, lidos da aba 2024R (se existir)."""
    codigos = {}
    if "2024R" in wb.sheetnames:
        for a, b in wb["2024R"].iter_rows(min_col=1, max_col=2, values_only=True):
            a, b = _texto(a), _texto(b)
            if a and b and re.match(r"^\d+(\.\d+)*$", a):
                codigos.setdefault(b, a)
                codigos.setdefault(b.lstrip("-> ").strip(), a)
    return codigos


def _conta(nome, cache, codigos):
    nome = _texto(nome)
    if not nome:
        return None
    if nome not in cache:
        conta, criada = Conta.objects.get_or_create(nome=nome, defaults={"codigo": codigos.get(nome, "")})
        if criada and "TRANSFER" in nome.upper():
            conta.nao_entra_no_dre = True
            conta.save(update_fields=["nao_entra_no_dre"])
        cache[nome] = conta
    return cache[nome]


@transaction.atomic
def importar(wb, ano, usuario=None, importar_regras=True, importar_lancamentos=True,
             substituir=True, coluna_valor=""):
    """Importa contas, regras e lançamentos. Retorna um dicionário com o resumo."""
    if importar_regras and "Regras" not in wb.sheetnames:
        raise ErroImportacao('A planilha não tem a aba "Regras".')
    meses_presentes = [m for m in ABAS_MESES if m in wb.sheetnames]
    if importar_lancamentos and not meses_presentes:
        raise ErroImportacao("Não encontrei nenhuma aba de mês (Jan, Fev, …).")

    motivo = "Importação da planilha"
    codigos = _codigos_contas(wb)
    cache = {c.nome: c for c in Conta.objects.all()}
    resumo = {"contas": 0, "regras": 0, "lancamentos": 0, "automaticos": 0, "manuais": 0, "sem_regra": 0}
    contas_antes = len(cache)

    if importar_regras:
        ws = wb["Regras"]
        for (nome,) in ws.iter_rows(min_row=5, min_col=8, max_col=8, values_only=True):
            if _texto(nome):
                _conta(nome, cache, codigos)
        if substituir:
            for r in Regra.objects.all():
                r._change_reason = motivo
                r.delete()
        novas = []
        for n, (palavra, tipo, secao, conta, obs) in enumerate(
            ws.iter_rows(min_row=5, min_col=1, max_col=5, values_only=True), start=5
        ):
            palavra = _texto(palavra)
            if not palavra or not _texto(conta):
                continue
            novas.append(Regra(
                palavra=palavra,
                tipo=TIPOS.get(_texto(tipo).upper(), "C"),
                secao=SECOES.get(_texto(secao).upper(), "A"),
                conta=_conta(conta, cache, codigos),
                ordem=n,
                observacao=_texto(obs)[:200],
            ))
        bulk_create_with_history(novas, Regra, default_user=usuario, default_change_reason=motivo)
        resumo["regras"] = len(novas)

    if importar_lancamentos:
        col_valor = column_index_from_string(coluna_valor.upper()) if coluna_valor else None
        if substituir:
            for lanc in Lancamento.objects.filter(ano=ano):
                lanc._change_reason = motivo
                lanc.delete()
        regras = regras_ativas()
        from .classificador import classificar

        novos = []
        for mes_num, mes_nome in MESES:
            if mes_nome not in wb.sheetnames:
                continue
            ws = wb[mes_nome]
            secao = "R"
            for linha in range(4, ws.max_row + 1):
                desc = _texto(ws.cell(linha, 2).value)
                if not desc:
                    continue
                if desc.upper() in CABECALHOS:
                    if desc.upper() == "DESPESAS":
                        secao = "D"
                    continue
                if desc.upper().startswith("TOTAL"):
                    continue
                conta_planilha = _texto(ws.cell(linha, 3).value)
                lanc = Lancamento(
                    ano=ano, mes=mes_num, secao=secao, descricao=desc[:255], linha_planilha=linha,
                    valor=_decimal(ws.cell(linha, col_valor).value) if col_valor else None,
                )
                regra = classificar(desc, secao, regras)
                if conta_planilha and (not regra or regra.conta.nome != conta_planilha):
                    # Alguém digitou por cima na planilha: mantém como correção manual.
                    lanc.conta = _conta(conta_planilha, cache, codigos)
                    lanc.classificacao = Lancamento.Classificacao.MANUAL
                    resumo["manuais"] += 1
                elif regra:
                    lanc.conta, lanc.regra = regra.conta, regra
                    lanc.classificacao = Lancamento.Classificacao.AUTO
                    resumo["automaticos"] += 1
                else:
                    lanc.classificacao = Lancamento.Classificacao.NENHUMA
                    resumo["sem_regra"] += 1
                novos.append(lanc)
        bulk_create_with_history(novos, Lancamento, batch_size=500, default_user=usuario,
                                 default_change_reason=motivo)
        resumo["lancamentos"] = len(novos)

    if importar_regras:
        # Regras novas valem também para os lançamentos de outros anos já cadastrados.
        from .classificador import reclassificar

        outros = Lancamento.objects.exclude(ano=ano) if importar_lancamentos else Lancamento.objects.all()
        resumo["reclassificados"] = reclassificar(outros, motivo=motivo, usuario=usuario)

    resumo["contas"] = len(cache) - contas_antes
    return resumo


def exportar(ano):
    """Gera um .xlsx com os lançamentos do ano, um resumo por conta e as regras."""
    wb = openpyxl.Workbook()
    negrito = Font(bold=True, color="FFFFFF")
    fundo = PatternFill("solid", fgColor="1F4E78")
    amarelo = PatternFill("solid", fgColor="FFF2CC")

    def cabecalho(ws, titulos, larguras):
        ws.append(titulos)
        for i, w in enumerate(larguras, 1):
            ws.cell(1, i).font, ws.cell(1, i).fill = negrito, fundo
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"

    ws = wb.active
    ws.title = f"Lançamentos {ano}"
    cabecalho(ws, ["Mês", "Seção", "Descrição", "Valor", "Plano de Contas", "Classificação", "Regra"],
              [7, 10, 45, 14, 36, 14, 24])
    for l in Lancamento.objects.filter(ano=ano).select_related("conta", "regra"):
        ws.append([l.get_mes_display(), l.get_secao_display(), l.descricao,
                   float(l.valor) if l.valor is not None else None,
                   l.conta.nome if l.conta else "", l.get_classificacao_display(),
                   l.regra.palavra if l.regra else ""])
        if not l.conta:
            for c in ws[ws.max_row]:
                c.fill = amarelo
        ws.cell(ws.max_row, 4).number_format = "#,##0.00"

    ws = wb.create_sheet("Regras")
    cabecalho(ws, ["Palavra-chave", "Tipo de busca", "Seção", "Plano de Contas", "Observação", "Ativa"],
              [26, 15, 11, 36, 40, 8])
    for r in Regra.objects.select_related("conta").order_by("ordem", "id"):
        ws.append([r.palavra, r.get_tipo_display(), r.get_secao_display(), r.conta.nome, r.observacao,
                   "Sim" if r.ativa else "Não"])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
