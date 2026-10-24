from django.db import migrations

BANCOS = [
    ("E", "SICOOB LM", "Sicoob"),
    ("F", "SICOOB RM", "Sicoob"),
    ("G", "ITAÚ", "Itaú"),
    ("H", "BRADESCO", "Bradesco"),
    ("I", "SICOOB RF", "Sicoob"),
    ("J", "LM SIST CX ECON", ""),
    ("K", "RM", ""),
]


def criar(apps, schema_editor):
    Banco = apps.get_model("financas", "Banco")
    for coluna, nome, instituicao in BANCOS:
        Banco.objects.get_or_create(nome=nome, defaults={"coluna_planilha": coluna, "banco": instituicao})


class Migration(migrations.Migration):
    dependencies = [("financas", "0005_bancos")]
    operations = [migrations.RunPython(criar, migrations.RunPython.noop)]
