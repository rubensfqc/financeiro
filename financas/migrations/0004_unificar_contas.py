from django.db import migrations


def unificar(apps, schema_editor):
    """Junta as contas criadas por importações anteriores com a conta do plano de mesmo nome."""
    from financas.contas import unificar as unificar_contas
    from financas.plano_contas import carregar

    Conta = apps.get_model("financas", "Conta")
    carregar(Conta)
    unificar_contas(Conta, apps.get_model("financas", "Lancamento"), apps.get_model("financas", "Regra"),
                    historico=False)


class Migration(migrations.Migration):
    dependencies = [("financas", "0003_rotulo_plano_de_contas")]
    operations = [migrations.RunPython(unificar, migrations.RunPython.noop)]
