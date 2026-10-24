"""Gera lançamentos fictícios com valores para ver o dashboard funcionando.

Os lançamentos ficam marcados com "[DEMO]" na descrição e são removidos com --apagar.
"""
import random
from decimal import Decimal

from django.core.management.base import BaseCommand, CommandError
from django.db import transaction
from django.utils import timezone

from financas.models import Conta, Lancamento
from financas.plano_contas import carregar

PREFIXO = "[DEMO]"

# código: (fração da receita ou valor fixo mensal, nº de lançamentos por mês)
VARIAVEIS = {
    "3.1": (0.21, 6), "3.2": (0.11, 5), "3.3": (0.03, 8), "3.4": (0.04, 6), "3.5": (0.03, 2),
    "3.6": (0.005, 1), "3.7": (0.13, 10), "3.8": (0.045, 1), "3.9": (0.03, 1), "3.10": (0.01, 1),
    "3.11": (0.03, 4), "3.12": (0.01, 2),
}
FIXOS = {
    "5.1": (62000, 4), "5.2": (5000, 1), "5.3": (13500, 1), "5.6": (3200, 1), "5.7": (2500, 1),
    "5.9": (9000, 1), "5.11": (11000, 1), "5.12": (900, 1), "5.15": (800, 1), "5.17": (6500, 3),
    "5.18": (1200, 2), "5.20": (4500, 1), "5.24": (2800, 3), "5.25": (25000, 1), "5.26": (5500, 1),
    "5.27": (1800, 1), "5.28": (650, 1), "5.30": (2000, 2),
}
RECEITAS = {"1.1": (0.10, 8), "1.2": (0.12, 10), "1.3": (0.45, 15), "1.5": (0.30, 25), "1.6": (0.03, 2)}
SAZONAL = [0.80, 0.85, 1.00, 1.05, 1.10, 1.05, 1.10, 1.15, 1.10, 1.05, 0.95, 0.80]


class Command(BaseCommand):
    help = "Cria (ou apaga, com --apagar) lançamentos fictícios para testar o dashboard."

    def add_arguments(self, parser):
        parser.add_argument("--anos", type=int, nargs="+", default=None)
        parser.add_argument("--apagar", action="store_true")
        parser.add_argument("--forcar", action="store_true", help="Gera mesmo se já houver lançamentos reais.")

    @transaction.atomic
    def handle(self, *args, **o):
        demo = Lancamento.objects.filter(descricao__startswith=PREFIXO)
        if o["apagar"]:
            n = demo.count()
            demo.delete()
            self.stdout.write(self.style.SUCCESS(f"{n} lançamentos de exemplo apagados."))
            return
        if Lancamento.objects.exclude(descricao__startswith=PREFIXO).exists() and not o["forcar"]:
            raise CommandError("Já existem lançamentos reais. Use --forcar se quiser misturar dados de exemplo.")
        demo.delete()
        carregar(Conta)
        contas = {c.codigo: c for c in Conta.objects.exclude(codigo="")}
        hoje = timezone.localdate()
        anos = o["anos"] or [hoje.year - 1, hoje.year]
        rnd = random.Random(42)
        novos = []
        for ano in anos:
            ultimo_mes = hoje.month - 1 if ano == hoje.year else 12
            for mes in range(1, max(ultimo_mes, 0) + 1):
                idx = (ano - anos[0]) * 12 + mes
                receita = 640000 * SAZONAL[mes - 1] * (1 + 0.006 * idx) * rnd.uniform(0.92, 1.08)
                aco_extra = 1 + (0.25 if (ano, mes) >= (anos[-1], 4) else 0)  # alta do aço no 2º tri
                itens = []
                for cod, (frac, n) in RECEITAS.items():
                    itens.append(("R", cod, receita * frac, n))
                for cod, (frac, n) in VARIAVEIS.items():
                    fator = aco_extra if cod == "3.1" else 1
                    itens.append(("D", cod, receita * frac * fator * rnd.uniform(0.85, 1.15), n))
                for cod, (valor, n) in FIXOS.items():
                    extra = 2.0 if cod == "5.1" and mes == 12 else 1  # 13º
                    itens.append(("D", cod, valor * extra * (1 + 0.004 * idx) * rnd.uniform(0.95, 1.05), n))
                if mes in (3, 9):
                    itens.append(("D", "7.2", 35000, 1))
                    itens.append(("D", "7.4", 6000, 1))
                for secao, cod, total, n in itens:
                    pesos = [rnd.uniform(0.5, 1.5) for _ in range(n)]
                    for i, p in enumerate(pesos, 1):
                        valor = Decimal(total * p / sum(pesos)).quantize(Decimal("0.01"))
                        novos.append(Lancamento(
                            ano=ano, mes=mes, secao=secao, valor=valor, conta=contas[cod],
                            classificacao=Lancamento.Classificacao.MANUAL,
                            descricao=f"{PREFIXO} {contas[cod].nome.lstrip('-> ')} {i}",
                        ))
        Lancamento.objects.bulk_create(novos, batch_size=1000)
        self.stdout.write(self.style.SUCCESS(f"{len(novos)} lançamentos de exemplo criados ({', '.join(map(str, anos))})."))
