"""Importação e exportação da planilha Gestão Financeira (.xlsx).

Estrutura esperada (a mesma da planilha com classificação automática):
- Aba "Regras": cabeçalho na linha 4, regras a partir da linha 5
  (A Palavra-chave | B Tipo de busca | C Seção | D Plano de Contas | E Observação),
  e a lista de contas válidas na coluna H a partir da linha 5.
- Abas de mês Jan…Dez: descrição na coluna B, plano de contas na coluna C,
  valor na coluna D, dados a partir da linha 4. Receitas acima da linha
  "DESPESAS", despesas abaixo dela.

O "Plano de Contas" da planilha é casado com as contas do app pelo nome
normalizado (veja contas.py), então pequenas diferenças de grafia não criam
contas duplicadas.
"""
import ast
import operator
import io
import re
from decimal import Decimal, InvalidOperation

import openpyxl
from django.db import transaction
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import column_index_from_string, get_column_letter
from simple_history.utils import bulk_create_with_history

from .classificador import regras_ativas
from .contas import Resolvedor, unificar
from .models import MESES, Conta, Lancamento, Regra

ABAS_MESES = [nome for _, nome in MESES]
TIPOS = {"CONTÉM": "C", "CONTEM": "C", "COMEÇA COM": "I", "COMECA COM": "I"}
SECOES = {"RECEITA": "R", "DESPESA": "D", "AMBAS": "A"}
CABECALHOS = {"DESPESAS", "RECEITAS", "DESCRIÇÃO", "DESCRICAO"}
COLUNA_VALOR_PADRAO = "D"


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
        wb = openpyxl.load_workbook(buf, data_only=True)
        buf.seek(0)
        # Também as fórmulas, para quando o arquivo não guardou o valor calculado
        # (acontece com planilhas salvas por outros programas).
        wb.formulas = openpyxl.load_workbook(buf, data_only=False)
        return wb
    except Exception as exc:
        raise ErroImportacao(f"Não consegui abrir o arquivo como .xlsx ({exc}).") from exc


_OPS = {ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul, ast.Div: operator.truediv}


def _avaliar_formula(formula):
    """Calcula fórmulas só de números e + - * / ( ), como '=1500+320,50'. Outras → None."""
    if not isinstance(formula, str) or not formula.startswith("="):
        return None
    expr = formula[1:].strip().replace(",", ".")
    if not re.fullmatch(r"[\d\s.+\-*/()]+", expr):
        return None

    def ev(n):
        if isinstance(n, ast.Expression):
            return ev(n.body)
        if isinstance(n, ast.Constant) and isinstance(n.value, (int, float)):
            return Decimal(str(n.value))
        if isinstance(n, ast.BinOp) and type(n.op) in _OPS:
            return _OPS[type(n.op)](ev(n.left), ev(n.right))
        if isinstance(n, ast.UnaryOp) and isinstance(n.op, (ast.USub, ast.UAdd)):
            return -ev(n.operand) if isinstance(n.op, ast.USub) else ev(n.operand)
        raise ValueError
    try:
        return ev(ast.parse(expr, mode="eval"))
    except Exception:
        return None


def _valor_celula(wb, aba, linha, coluna):
    v = _decimal(wb[aba].cell(linha, coluna).value)
    if v is None and getattr(wb, "formulas", None) is not None:
        bruto = _avaliar_formula(wb.formulas[aba].cell(linha, coluna).value)
        if bruto is not None:
            v = bruto.quantize(Decimal("0.01"))
    return v


def _texto(v):
    return str(v).strip() if v is not None else ""


def _decimal(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float, Decimal)):
        return Decimal(str(v)).quantize(Decimal("0.01"))
    texto = str(v).strip()
    negativo = texto.startswith("(") and texto.endswith(")")
    s = re.sub(r"[^\d,.-]", "", texto)
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(\.\d{3})+", s):
        s = s.replace(".", "")  # 1.500 = mil e quinhentos
    if negativo and not s.startswith("-"):
        s = "-" + s
    try:
        return Decimal(s).quantize(Decimal("0.01"))
    except InvalidOperation:
        return None





def _linhas_mes(wb, mes_nome):
    """(linha, seção, descrição, plano de contas) de cada lançamento da aba do mês."""
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
        yield linha, secao, desc, _texto(ws.cell(linha, 3).value)


@transaction.atomic
def atualizar_valores(wb, ano, usuario=None, coluna_valor=COLUNA_VALOR_PADRAO):
    """Só preenche/atualiza o valor dos lançamentos já importados (mesmo mês, linha e descrição).

    Não mexe em contas, regras nem nas correções feitas no app.
    """
    from simple_history.utils import bulk_update_with_history

    col = column_index_from_string((coluna_valor or COLUNA_VALOR_PADRAO).upper())
    existentes = {(l.mes, l.linha_planilha, l.descricao): l
                  for l in Lancamento.objects.filter(ano=ano, linha_planilha__isnull=False)}
    alterados, nao_achados, iguais = [], 0, 0
    for mes_num, mes_nome in MESES:
        if mes_nome not in wb.sheetnames:
            continue
        for linha, _secao, desc, _conta in _linhas_mes(wb, mes_nome):
            lanc = existentes.get((mes_num, linha, desc[:255]))
            if lanc is None:
                nao_achados += 1
                continue
            valor = _valor_celula(wb, mes_nome, linha, col)
            if valor == lanc.valor:
                iguais += 1
                continue
            lanc.valor = valor
            alterados.append(lanc)
    motivo = f"Valores atualizados da planilha (coluna {(coluna_valor or COLUNA_VALOR_PADRAO).upper()})"
    if alterados:
        bulk_update_with_history(alterados, Lancamento, ["valor"], batch_size=500,
                                 default_user=usuario, default_change_reason=motivo)
    return {"valores_atualizados": len(alterados), "valores_iguais": iguais, "nao_encontrados": nao_achados,
            "contas_unificadas": unificar(Conta, Lancamento, Regra, usuario=usuario)}


@transaction.atomic
def importar(wb, ano, usuario=None, importar_regras=True, importar_lancamentos=True,
             substituir=True, coluna_valor=COLUNA_VALOR_PADRAO):
    """Importa contas, regras e lançamentos. Retorna um dicionário com o resumo."""
    if importar_regras and "Regras" not in wb.sheetnames:
        raise ErroImportacao('A planilha não tem a aba "Regras".')
    meses_presentes = [m for m in ABAS_MESES if m in wb.sheetnames]
    if importar_lancamentos and not meses_presentes:
        raise ErroImportacao("Não encontrei nenhuma aba de mês (Jan, Fev, …).")

    motivo = "Importação da planilha"
    # Junta contas duplicadas de importações anteriores antes de começar.
    resumo = {"contas": 0, "regras": 0, "lancamentos": 0, "automaticos": 0, "manuais": 0, "sem_regra": 0,
              "com_valor": 0, "contas_unificadas": unificar(Conta, Lancamento, Regra, usuario=usuario)}
    # Só o plano de contas oficial (tabela da empresa) define códigos: um nome que não
    # casar com ele entra sem código e aparece em "Qualidade dos dados" para ser mesclado.
    conta_de = Resolvedor(Conta)

    if importar_regras:
        ws = wb["Regras"]
        for (nome,) in ws.iter_rows(min_row=5, min_col=8, max_col=8, values_only=True):
            if _texto(nome):
                conta_de(nome)
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
                conta=conta_de(conta),
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
            for linha, secao, desc, conta_planilha in _linhas_mes(wb, mes_nome):
                lanc = Lancamento(
                    ano=ano, mes=mes_num, secao=secao, descricao=desc[:255], linha_planilha=linha,
                    valor=_valor_celula(wb, mes_nome, linha, col_valor) if col_valor else None,
                )
                if lanc.valor is not None:
                    resumo["com_valor"] += 1
                regra = classificar(desc, secao, regras)
                conta = conta_de(conta_planilha)
                if conta and (not regra or regra.conta_id != conta.id):
                    # Alguém digitou por cima na planilha: mantém como correção manual.
                    lanc.conta = conta
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

    resumo["contas"] = conta_de.criadas
    resumo["contas_sem_codigo"] = list(
        Conta.objects.filter(codigo="", nao_entra_no_dre=False).values_list("nome", flat=True))
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
