from getpass import getpass

from django.contrib.auth import get_user_model
from django.core.management.base import BaseCommand, CommandError

from financas import planilha


class Command(BaseCommand):
    help = ("Importa regras, contas e lançamentos da planilha Gestão Financeira (.xlsx) "
            "ou de uma planilha exportada pelo app.")

    def add_arguments(self, parser):
        parser.add_argument("arquivo")
        parser.add_argument("--ano", type=int, default=None,
                            help="Ano dos lançamentos (obrigatório na planilha modelo; na exportada vem da aba).")
        parser.add_argument("--senha", default=None, help="Senha da planilha (se omitida e necessária, é pedida).")
        parser.add_argument("--coluna-valor", default="D", help="Letra da coluna com o valor (padrão: D).")
        parser.add_argument("--apenas-valores", action="store_true",
                            help="Só atualiza os valores dos lançamentos já importados (não mexe em contas/regras).")
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
            if o["ano"] is None:
                o["ano"] = planilha.ano_da_planilha(wb) if planilha.formato(wb) == "exportada" else None
                if o["ano"] is None:
                    raise CommandError("Informe --ano.")
            if o["apenas_valores"]:
                r = planilha.atualizar_valores(wb, o["ano"], usuario=usuario, coluna_valor=o["coluna_valor"])
                self.stdout.write(self.style.SUCCESS(
                    f"Valores atualizados: {r['valores_atualizados']} | iguais: {r['valores_iguais']} | "
                    f"linhas sem lançamento no app: {r['nao_encontrados']}"))
                return
            r = planilha.importar(
                wb, o["ano"], usuario=usuario, importar_regras=not o["sem_regras"],
                importar_lancamentos=not o["sem_lancamentos"], coluna_valor=o["coluna_valor"],
            )
        except planilha.ErroImportacao as exc:
            raise CommandError(str(exc)) from exc
        if r.get("contas_sem_codigo"):
            self.stdout.write(self.style.WARNING(
                "Nomes de plano de contas fora do plano (sem código): " + ", ".join(r["contas_sem_codigo"])))
        self.stdout.write(f"Formato: {'planilha exportada pelo app' if r['formato'] == 'exportada' else 'planilha modelo'} | ano {r['ano']}")
        self.stdout.write(self.style.SUCCESS(
            f"Contas novas: {r['contas']} | Regras: {r['regras']} | Lançamentos: {r['lancamentos']} "
            f"(auto {r['automaticos']}, manual {r['manuais']}, sem regra {r['sem_regra']}, com valor {r['com_valor']}) | "
            f"contas unificadas: {r['contas_unificadas']}"
        ))
