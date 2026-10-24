from django.core.management.base import BaseCommand

from financas.models import Conta
from financas.plano_contas import carregar


class Command(BaseCommand):
    help = "Cria as contas do plano de contas que faltam e acerta os códigos das existentes."

    def handle(self, *args, **o):
        criadas, atualizadas = carregar(Conta)
        sem = Conta.objects.filter(codigo="", nao_entra_no_dre=False).values_list("nome", flat=True)
        self.stdout.write(self.style.SUCCESS(f"Contas criadas: {criadas} | códigos atualizados: {atualizadas}"))
        if sem:
            self.stdout.write(self.style.WARNING("Contas sem código (ficam fora do dashboard): " + ", ".join(sem)))
