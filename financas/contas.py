"""Casamento entre o "Plano de Contas" da planilha e as contas do app.

A planilha pode trazer o mesmo item escrito de jeitos diferentes
("-> Compras de Aço", "->Compras de Aço", "Compras de aço ", "3.1 -> Compras de Aço").
Todas essas formas têm a mesma chave e viram uma conta só: a do plano de contas,
que tem código e por isso entra no dashboard.
"""
import re
import unicodedata

_PREFIXO = re.compile(r"^\s*(\d+(\.\d+)*)?\s*[-–—=→]*\s*>?\s*")


def chave(nome):
    """Forma normalizada do nome: sem código/seta na frente, sem acentos, minúsculas."""
    s = unicodedata.normalize("NFKD", str(nome or "").replace("\xa0", " "))
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = _PREFIXO.sub("", s)
    return " ".join(s.casefold().split())


def codigo_no_nome(nome):
    m = re.match(r"^\s*(\d+\.\d+)\b", str(nome or ""))
    return m.group(1) if m else None


class Resolvedor:
    """Acha (ou cria) a conta de um nome vindo da planilha."""

    def __init__(self, conta_model, codigos_extra=None):
        self.Conta = conta_model
        self.por_chave = {}
        self.por_codigo = {}
        self.criadas = 0
        self.codigos_extra = {chave(k): v for k, v in (codigos_extra or {}).items()}
        # Contas com código têm preferência quando duas têm a mesma chave.
        for c in conta_model.objects.order_by("-codigo", "id"):
            self.por_chave.setdefault(chave(c.nome), c)
            if c.codigo:
                self.por_codigo.setdefault(c.codigo, c)

    def __call__(self, nome):
        nome = str(nome or "").strip()
        if not nome:
            return None
        k = chave(nome)
        if not k:
            return None
        conta = self.por_chave.get(k)
        if conta is None:
            cod = codigo_no_nome(nome)
            conta = self.por_codigo.get(cod) if cod else None
        if conta is None:
            cod = codigo_no_nome(nome) or self.codigos_extra.get(k, "")
            if cod and cod in self.por_codigo:
                cod = ""  # código já usado por outra conta
            conta = self.Conta.objects.create(
                nome=nome, codigo=cod, nao_entra_no_dre="transfer" in k,
            )
            self.criadas += 1
            if cod:
                self.por_codigo[cod] = conta
        self.por_chave[k] = conta
        return conta


def mesclar(origem, destino, Lancamento, Regra, usuario=None, historico=True):
    """Move lançamentos e regras de `origem` para `destino` e apaga `origem`."""
    motivo = f'Contas unificadas: "{origem.nome}" → "{destino.nome}"'
    lancs = list(Lancamento.objects.filter(conta=origem))
    regras = list(Regra.objects.filter(conta=origem))
    for obj in lancs + regras:
        obj.conta = destino
        obj._change_reason = motivo
    if historico:
        from simple_history.utils import bulk_update_with_history

        if lancs:
            bulk_update_with_history(lancs, Lancamento, ["conta"], batch_size=500,
                                     default_user=usuario, default_change_reason=motivo)
        if regras:
            bulk_update_with_history(regras, Regra, ["conta"], default_user=usuario,
                                     default_change_reason=motivo)
        origem._change_reason = motivo
    else:
        Lancamento.objects.filter(conta=origem).update(conta=destino)
        Regra.objects.filter(conta=origem).update(conta=destino)
    if origem.nao_entra_no_dre and not destino.nao_entra_no_dre and not destino.codigo:
        destino.nao_entra_no_dre = True
        destino.save(update_fields=["nao_entra_no_dre"])
    origem.delete()
    return len(lancs), len(regras)


def unificar(Conta, Lancamento, Regra, usuario=None, historico=True):
    """Junta as contas que são a mesma (mesma chave). Fica a que tem código."""
    grupos = {}
    for c in Conta.objects.order_by("id"):
        grupos.setdefault(chave(c.nome), []).append(c)
    unidas = 0
    for contas in grupos.values():
        if len(contas) < 2:
            continue
        contas.sort(key=lambda c: (not c.codigo, c.id))
        destino = contas[0]
        for origem in contas[1:]:
            mesclar(origem, destino, Lancamento, Regra, usuario=usuario, historico=historico)
            unidas += 1
    return unidas
