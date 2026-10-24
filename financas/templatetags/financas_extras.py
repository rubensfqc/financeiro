from decimal import Decimal

from django import template

register = template.Library()


def _milhar(n, casas=0):
    n = round(Decimal(n), casas)
    if n == 0:
        n = abs(n)  # evita "-0"
    s = f"{abs(n):,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("-" if n < 0 else "") + s


@register.filter
def brl(valor, casas=0):
    """1234567.8 -> 'R$ 1.234.568'."""
    if valor is None or valor == "":
        return "—"
    return "R$ " + _milhar(Decimal(valor), int(casas))


@register.filter
def pct(valor, casas=1):
    if valor is None:
        return "—"
    return _milhar(Decimal(valor), int(casas)) + "%"


@register.filter
def delta(info):
    """Formata a variação de um cartão: '+5,2%' ou '-1,3 p.p.'."""
    if not info:
        return ""
    d = Decimal(info["delta"])
    sinal = "+" if d > 0 else ("−" if d < 0 else "")
    return f"{sinal}{_milhar(abs(d), 1)}{' p.p.' if info['unidade'] == 'p.p.' else '%'}"


@register.filter
def valor_cartao(cartao):
    return pct(cartao["valor"]) if cartao["formato"] == "%" else brl(cartao["valor"])


@register.filter
def pct_de(parte, todo):
    """{{ parte|pct_de:todo }} -> '12,3%'."""
    if not todo:
        return "—"
    return pct(Decimal(parte) / Decimal(todo) * 100)
