"""Plano de contas da empresa (mesmo da aba 2024R / DRE da planilha).

O grupo de cada conta é o número antes do ponto no código (3.1 → grupo 3).
"""

GRUPOS = {
    1: "Receitas de vendas",
    2: "Receitas não operacionais",
    3: "Custos variáveis",
    5: "Custos fixos",
    7: "Despesas não operacionais",
}

PLANO = [
    ("1.1", "-> Rec Vendas a Vista"),
    ("1.2", "-> Rec Vendas Cartão Cred/Deb"),
    ("1.3", "-> Rec Vendas Faturada"),
    ("1.4", "-> Rec Vendas Cheque Pré"),
    ("1.5", "-> Rec PIX"),
    ("1.6", "-> Outros Rec Vendas"),
    ("2.1", "-> Aportes de Capital Sócios"),
    ("2.2", "-> Receb Empréstimos"),
    ("2.3", "-> Rec Venda de Ativos"),
    ("2.4", "-> Outras Rec Não Operacionais"),
    ("3.1", "-> Compras de Aço"),
    ("3.2", "-> Compras de Cimento"),
    ("3.3", "-> Compras de Combustível"),
    ("3.4", "-> Fretes em Geral"),
    ("3.5", "-> Comissão Vendas"),
    ("3.6", "-> Comissão Adm Cartões"),
    ("3.7", "-> Compra Mercadorias Diversas"),
    ("3.8", "-> Impostos Federais"),
    ("3.9", "-> Impostos ICMS"),
    ("3.10", "-> Impostos ISS"),
    ("3.11", "-> MO Avulsa"),
    ("3.12", "-> Outras Despesas Variáveis"),
    ("5.1", "-> Salários e Vales"),
    ("5.2", "-> Encargos FGTS"),
    ("5.3", "-> Encargos INSS"),
    ("5.4", "-> Encargos 13o Salário"),
    ("5.5", "-> Rescisões Trabalhistas"),
    ("5.6", "-> Vale Transporte"),
    ("5.7", "-> Benefícios Diversos"),
    ("5.8", "-> Outras Despesas c/Pessoal"),
    ("5.9", "-> Aluguel e Condomínio"),
    ("5.10", "-> Taxas Diversas Imóvel"),
    ("5.11", "-> Energia Elétrica"),
    ("5.12", "-> Abastecimento de Água"),
    ("5.13", "-> Abastecimento de Gás"),
    ("5.14", "-> Telefonia Fixa"),
    ("5.15", "-> Telefonia Móvel e Internet"),
    ("5.16", "-> Provedores Internet+Site"),
    ("5.17", "-> Manutenção Geral"),
    ("5.18", "-> Material Escritório e Limpeza"),
    ("5.19", "-> Despesas de Locomoção"),
    ("5.20", "-> Contador e Advogados"),
    ("5.21", "-> Cópias Correios e Cartórios"),
    ("5.22", "-> Sindicatos e Assoc Classe"),
    ("5.23", "-> Assinatura Jornais e Revistas"),
    ("5.24", "-> Lanches e Refeições"),
    ("5.25", "-> Prolabore Sócios"),
    ("5.26", "-> Encargos s/Prolabore"),
    ("5.27", "-> Seguros em Geral"),
    ("5.28", "-> Tarifas Bancárias"),
    ("5.30", "-> Outras Despesas Fixas"),
    ("7.1", "-> Distribuição de Lucros"),
    ("7.2", "-> Pagamento Empréstimos"),
    ("7.3", "-> Investimentos em Geral"),
    ("7.4", "-> Juros"),
    ("7.5", "-> Outras Despesas Não Operc"),
]

# Indicadores específicos (por código de conta).
MATERIA_PRIMA = {"3.1": "Aço", "3.2": "Cimento", "3.7": "Mercadorias diversas"}
LOGISTICA = ("3.3", "3.4")                       # combustível + fretes
COMISSOES = ("3.5", "3.6")
IMPOSTOS = ("3.8", "3.9", "3.10")
MAO_DE_OBRA = ("3.11", "5.1", "5.2", "5.3", "5.4", "5.5", "5.6", "5.7", "5.8")
PROLABORE = ("5.25", "5.26")
FORMAS_RECEBIMENTO = {
    "1.1": "À vista", "1.2": "Cartão", "1.3": "Faturada", "1.4": "Cheque pré", "1.5": "PIX", "1.6": "Outros",
}


def grupo_do_codigo(codigo):
    try:
        return int(str(codigo).split(".")[0])
    except (TypeError, ValueError):
        return None


def carregar(conta_model):
    """Cria as contas que faltam e acerta o código das existentes (pelo nome)."""
    criadas = atualizadas = 0
    for codigo, nome in PLANO:
        conta = conta_model.objects.filter(nome=nome).first()
        if conta is None:
            conta_model.objects.create(codigo=codigo, nome=nome)
            criadas += 1
        elif conta.codigo != codigo:
            conta.codigo = codigo
            conta.save(update_fields=["codigo"])
            atualizadas += 1
    return criadas, atualizadas
