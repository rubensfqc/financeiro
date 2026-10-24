from getpass import getpass

from django.contrib.auth import get_user_model
from django.core.management.base import BaseCommand, CommandError

from financas import planilha


class Command(BaseCommand):
    help = "Importa regras, contas e lançamentos da planilha Gestão Financeira (.xlsx)."

    def add_arguments(self, parser):
        parser.add_argument("arquivo")
        parser.add_argument("--ano", type=int, required=True)
        parser.add_argument("--senha", default=None, help="Senha da planilha (se omitida e necessária, é pedida).")
        parser.add_argument("--coluna-valor", default="", help="Letra da coluna com o valor (ex.: D).")
        parser.add_argument("--sem-regras", action="store_true", help="Não importa a aba Regras.")
        parser.add_argument("--sem-lancamentos", action="store_true", help="Não importa as abas de mês.")
        parser.add_argument("--usuario", default=None, help="Usuário que aparece no histórico.")

    def handle(self, *args, **o):
        usuario = None
        if o["usuario"]:
            usuario = get_user_model().objects.filter(username=o["usuario"]).first()
            if not usuario:
                raise CommandError(f"Usuário {o['usuario']} não existe.")
        senha = o["senha"]
        try:
            try:
                wb = planilha.abrir_planilha(o["arquivo"], senha or "")
            except planilha.ErroImportacao as exc:
                if senha is None and "tem senha" in str(exc):
                    wb = planilha.abrir_planilha(o["arquivo"], getpass("Senha da planilha: "))
                else:
                    raise
            r = planilha.importar(
                wb, o["ano"], usuario=usuario, importar_regras=not o["sem_regras"],
                importar_lancamentos=not o["sem_lancamentos"], coluna_valor=o["coluna_valor"],
            )
        except planilha.ErroImportacao as exc:
            raise CommandError(str(exc)) from exc
        self.stdout.write(self.style.SUCCESS(
            f"Contas novas: {r['contas']} | Regras: {r['regras']} | Lançamentos: {r['lancamentos']} "
            f"(auto {r['automaticos']}, manual {r['manuais']}, sem regra {r['sem_regra']})"
        ))
