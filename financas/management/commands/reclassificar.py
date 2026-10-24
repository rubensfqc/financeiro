from django.core.management.base import BaseCommand

from financas.classificador import reclassificar


class Command(BaseCommand):
    help = "Reaplica as regras a todos os lançamentos não corrigidos à mão."

    def handle(self, *args, **o):
        n = reclassificar(motivo="Reclassificação (linha de comando)")
        self.stdout.write(self.style.SUCCESS(f"{n} lançamento(s) mudaram de conta."))
