from django.db import migrations


def carregar(apps, schema_editor):
    from financas.plano_contas import carregar as carregar_plano

    carregar_plano(apps.get_model("financas", "Conta"))
    Conta = apps.get_model("financas", "Conta")
    Conta.objects.filter(nome__icontains="transfer").update(nao_entra_no_dre=True)


class Migration(migrations.Migration):
    dependencies = [("financas", "0001_initial")]
    operations = [migrations.RunPython(carregar, migrations.RunPython.noop)]
