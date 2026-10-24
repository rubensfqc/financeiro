"""Importação e exportação da planilha Gestão Financeira (.xlsx).

Estrutura esperada (a mesma da planilha com classificação automática):
- Aba "Regras": cabeçalho na linha 4, regras a partir da linha 5
  (A Palavra-chave | B Tipo de busca | C Seção | D Plano de Contas | E Observação),
  e a lista de contas válidas na coluna H a partir da linha 5.
- Abas de mês Jan…Dez: descrição na coluna B, plano de contas na coluna C,
  valor na coluna D e uma coluna por banco (E…K, configurável no cadastro de
  bancos), dados a partir da linha 4. Receitas acima da linha
  "DESPESAS", despesas abaixo dela.

O "Plano de Contas" da planilha é casado com as contas do app pelo nome
normalizado (veja contas.py), então pequenas diferenças de grafia não criam
contas duplicadas.
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


class _Formula:
    """Calcula as fórmulas comuns da planilha a partir das próprias células.

    Usado quando o arquivo não guardou o resultado calculado (por exemplo, depois
    de salvo por outro programa). Entende números, referências (E5, $E$5),
    intervalos (E5:K5), SOMA/SUM e + - * / ( ). Célula vazia vale 0, como no Excel.
    Qualquer outra coisa (outra aba, outras funções) → None.
    """

    TOKEN = re.compile(r"\s*(?:(?P<num>\d+(?:[.,]\d+)?)|(?P<rng>\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+)"
                       r"|(?P<ref>\$?[A-Z]{1,3}\$?\d+)|(?P<fn>SUM|SOMA)\s*\(|(?P<op>[-+*/();,]))", re.I)

    def __init__(self, wb, aba, profundidade=0):
        self.wb, self.aba, self.prof = wb, aba, profundidade

    def calcular(self, formula):
        if not isinstance(formula, str) or not formula.startswith("=") or self.prof > 8:
            return None
        texto = formula[1:].strip()
        self.toks, pos = [], 0
        while pos < len(texto):
            m = self.TOKEN.match(texto, pos)
            if not m or m.end() == pos:
                return None
            self.toks.append((m.lastgroup, m.group(m.lastgroup)))
            pos = m.end()
        self.i = 0
        try:
            v = self._expr()
            return v if self.i == len(self.toks) else None
        except (ValueError, ArithmeticError, IndexError):
            return None

    def _ver(self):
        return self.toks[self.i] if self.i < len(self.toks) else (None, None)

    def _op(self, *ops):
        t, v = self._ver()
        if t == "op" and v in ops:
            self.i += 1
            return v
        return None

    def _expr(self):
        v = self._termo()
        while (op := self._op("+", "-")):
            v = v + self._termo() if op == "+" else v - self._termo()
        return v

    def _termo(self):
        v = self._fator()
        while (op := self._op("*", "/")):
            d = self._fator()
            v = v * d if op == "*" else v / d
        return v

    def _fator(self):
        if self._op("-"):
            return -self._fator()
        if self._op("+"):
            return self._fator()
        if self._op("("):
            v = self._expr()
            if not self._op(")"):
                raise ValueError
            return v
        t, v = self._ver()
        self.i += 1
        if t == "num":
            return Decimal(v.replace(",", "."))
        if t == "ref":
            return self._celula(v)
        if t == "fn":
            total = Decimal(0)
            while not self._op(")"):
                t2, v2 = self._ver()
                if t2 == "rng":
                    self.i += 1
                    total += sum((self._celula(c.coordinate) for linha in self.wb[self.aba][v2.replace("$", "")]
                                  for c in linha), Decimal(0))
                else:
                    total += self._expr()
                self._op(",", ";")
            return total
        raise ValueError

    def _celula(self, ref):
        ref = ref.replace("$", "")
        v = _decimal(self.wb[self.aba][ref].value)
        if v is None and getattr(self.wb, "formulas", None) is not None:
            v = _Formula(self.wb, self.aba, self.prof + 1).calcular(self.wb.formulas[self.aba][ref].value)
        return v if v is not None else Decimal(0)


def _valor_celula(wb, aba, linha, coluna):
    """Valor numérico da célula: o calculado pelo Excel ou, se não houver, calculado aqui. Vazia → None."""
    v = _decimal(wb[aba].cell(linha, coluna).value)
    if v is None and getattr(wb, "formulas", None) is not None:
        bruto = _Formula(wb, aba).calcular(wb.formulas[aba].cell(linha, coluna).value)
        if bruto is not None:
            v = bruto.quantize(Decimal("0.01"))
    return v


def _colunas_bancos():
    """[(índice da coluna, banco)] dos bancos ativos com coluna definida."""
    from .models import Banco

    r = []
    for b in Banco.objects.filter(ativo=True).exclude(coluna_planilha=""):
        try:
            r.append((column_index_from_string(b.coluna_planilha.strip().upper()), b))
        except ValueError:
            pass
    return sorted(r, key=lambda x: x[0])


def _banco_da_linha(wb, aba, linha, colunas_bancos):
    """Banco cuja coluna tem valor diferente de zero; (banco, valor, quantos bancos com valor)."""
    achados = []
    for col, banco in colunas_bancos:
        v = _valor_celula(wb, aba, linha, col)
        if v is not None and v != 0:
            achados.append((banco, v))
    if not achados:
        return None, None, 0
    return achados[0][0], achados[0][1], len(achados)


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





def _valor_e_banco(wb, aba, linha, col_valor, colunas_bancos, resumo):
    """Valor (coluna D) e banco (coluna E–K com valor diferente de zero) de uma linha."""
    valor = _valor_celula(wb, aba, linha, col_valor) if col_valor else None
    banco, valor_banco, n = _banco_da_linha(wb, aba, linha, colunas_bancos)
    if valor is None and valor_banco is not None and n == 1:
        valor = valor_banco  # D vazia/ilegível, mas o valor está na coluna do banco
        resumo["valor_pela_coluna_do_banco"] = resumo.get("valor_pela_coluna_do_banco", 0) + 1
    if n > 1:
        resumo.setdefault("varios_bancos", []).append(f"{aba} linha {linha}")
    if valor is not None:
        resumo["com_valor"] = resumo.get("com_valor", 0) + 1
    if banco is not None:
        resumo["com_banco"] = resumo.get("com_banco", 0) + 1
    return valor, banco


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
    """Só preenche/atualiza valor e banco dos lançamentos já importados (mesmo mês, linha e descrição).

    Não mexe em plano de contas, regras nem nas correções feitas no app.
    """
    from simple_history.utils import bulk_update_with_history

    col = column_index_from_string((coluna_valor or COLUNA_VALOR_PADRAO).upper())
    existentes = {(l.mes, l.linha_planilha, l.descricao): l
                  for l in Lancamento.objects.filter(ano=ano, linha_planilha__isnull=False)}
    bancos = _colunas_bancos()
    resumo = {}
    alterados, nao_achados, iguais = [], 0, 0
    for mes_num, mes_nome in MESES:
        if mes_nome not in wb.sheetnames:
            continue
        for linha, _secao, desc, _conta in _linhas_mes(wb, mes_nome):
            lanc = existentes.get((mes_num, linha, desc[:255]))
            if lanc is None:
                nao_achados += 1
                continue
            valor, banco = _valor_e_banco(wb, mes_nome, linha, col, bancos, resumo)
            if (valor, banco.id if banco else None) == (lanc.valor, lanc.banco_id):
                iguais += 1
                continue
            lanc.valor, lanc.banco = valor, banco
            alterados.append(lanc)
    motivo = f"Valores e bancos atualizados da planilha (valor da coluna {(coluna_valor or COLUNA_VALOR_PADRAO).upper()})"
    if alterados:
        bulk_update_with_history(alterados, Lancamento, ["valor", "banco"], batch_size=500,
                                 default_user=usuario, default_change_reason=motivo)
    resumo.update({"valores_atualizados": len(alterados), "valores_iguais": iguais, "nao_encontrados": nao_achados,
                   "contas_unificadas": unificar(Conta, Lancamento, Regra, usuario=usuario)})
    return resumo


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
              "com_valor": 0, "com_banco": 0, "contas_unificadas": unificar(Conta, Lancamento, Regra, usuario=usuario)}
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
        bancos = _colunas_bancos()
        from .classificador import classificar

        novos = []
        for mes_num, mes_nome in MESES:
            if mes_nome not in wb.sheetnames:
                continue
            for linha, secao, desc, conta_planilha in _linhas_mes(wb, mes_nome):
                valor, banco = _valor_e_banco(wb, mes_nome, linha, col_valor, bancos, resumo)
                lanc = Lancamento(
                    ano=ano, mes=mes_num, secao=secao, descricao=desc[:255], linha_planilha=linha,
                    valor=valor, banco=banco,
                )
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
    cabecalho(ws, ["Mês", "Seção", "Descrição", "Valor", "Banco", "Plano de Contas", "Classificação", "Regra"],
              [7, 10, 45, 14, 18, 36, 14, 24])
    for l in Lancamento.objects.filter(ano=ano).select_related("conta", "regra", "banco"):
        ws.append([l.get_mes_display(), l.get_secao_display(), l.descricao,
                   float(l.valor) if l.valor is not None else None, l.banco.nome if l.banco else "",
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
