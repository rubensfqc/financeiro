"""Classificação automática por palavra-chave (mesma lógica da aba Regras).

- Maiúsculas/minúsculas não importam; acentos importam (cadastre AÇO e ACO).
- "Contém": a palavra-chave aparece no início de qualquer palavra da descrição
  (PEDRA casa com "PEDRA BRITADA" e "COMPRA PEDRAS", mas não com "EMPEDRADO").
- "Começa com": a descrição começa com a palavra-chave.
- A regra vale para Receita, Despesa ou Ambas.
- Se várias servirem, vence a de maior 'ordem'.
"""
import re

from django.db import transaction

_SEPARADORES = re.compile(r"[^\w]+")


def normalizar(texto):
    return " ".join(_SEPARADORES.sub(" ", (texto or "").upper()).split())


def regras_ativas():
    """Regras em ordem de prioridade decrescente, já normalizadas."""
    from .models import Regra

    regras = list(Regra.objects.filter(ativa=True).select_related("conta").order_by("-ordem", "-id"))
    for r in regras:
        r.chave = normalizar(r.palavra)
    return regras


def regra_serve(regra, descricao_norm, secao):
    if regra.secao != "A" and regra.secao != secao:
        return False
    if not regra.chave:
        return False
    if regra.tipo == "I":
        return descricao_norm.startswith(regra.chave)
    return (" " + descricao_norm).find(" " + regra.chave) >= 0


def classificar(descricao, secao, regras=None):
    """Retorna a Regra vencedora ou None."""
    if regras is None:
        regras = regras_ativas()
    d = normalizar(descricao)
    for regra in regras:  # já em ordem de prioridade
        if regra_serve(regra, d, secao):
            return regra
    return None


def reclassificar(qs=None, motivo="Reclassificação pelas regras", usuario=None):
    """Reaplica as regras a todos os lançamentos não corrigidos à mão.

    Só grava (e só gera histórico) o que realmente mudou.
    """
    from simple_history.utils import bulk_update_with_history

    from .models import Lancamento

    if qs is None:
        qs = Lancamento.objects.all()
    regras = regras_ativas()
    alterados = []
    for lanc in qs.exclude(classificacao=Lancamento.Classificacao.MANUAL).iterator(chunk_size=1000):
        if lanc.aplicar_regras(regras):
            lanc._change_reason = motivo
            alterados.append(lanc)
    if alterados:
        with transaction.atomic():
            bulk_update_with_history(
                alterados, Lancamento, ["conta", "regra", "classificacao"], batch_size=500,
                default_change_reason=motivo, default_user=usuario,
            )
    return len(alterados)
