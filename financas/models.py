from django.db import models
from simple_history.models import HistoricalRecords

MESES = [
    (1, "Jan"), (2, "Fev"), (3, "Mar"), (4, "Abr"), (5, "Mai"), (6, "Jun"),
    (7, "Jul"), (8, "Ago"), (9, "Set"), (10, "Out"), (11, "Nov"), (12, "Dez"),
]


class Secao(models.TextChoices):
    RECEITA = "R", "Receita"
    DESPESA = "D", "Despesa"


class Conta(models.Model):
    """Uma linha do plano de contas (ex.: '3.1 Compras de Aço')."""

    codigo = models.CharField("código", max_length=20, blank=True)
    nome = models.CharField(max_length=120, unique=True)
    nao_entra_no_dre = models.BooleanField(
        "fora do DRE", default=False,
        help_text="Ex.: transferência entre contas da própria empresa.",
    )
    ativa = models.BooleanField(default=True)

    history = HistoricalRecords()

    class Meta:
        ordering = ["codigo", "nome"]
        verbose_name = "conta do plano de contas"
        verbose_name_plural = "plano de contas"

    def __str__(self):
        return self.nome


class Regra(models.Model):
    """Palavra-chave que leva uma descrição a uma conta.

    Mesma lógica da aba Regras da planilha: se mais de uma regra servir,
    vale a de maior 'ordem' (na planilha, a que está mais abaixo).
    """

    class Tipo(models.TextChoices):
        CONTEM = "C", "Contém"
        COMECA = "I", "Começa com"

    class SecaoRegra(models.TextChoices):
        RECEITA = "R", "Receita"
        DESPESA = "D", "Despesa"
        AMBAS = "A", "Ambas"

    palavra = models.CharField("palavra-chave", max_length=80)
    tipo = models.CharField("tipo de busca", max_length=1, choices=Tipo.choices, default=Tipo.CONTEM)
    secao = models.CharField("seção", max_length=1, choices=SecaoRegra.choices, default=SecaoRegra.AMBAS)
    conta = models.ForeignKey(Conta, on_delete=models.PROTECT, related_name="regras", verbose_name="plano de contas")
    ordem = models.PositiveIntegerField(
        default=0, db_index=True,
        help_text="Maior número = maior prioridade (exceções ficam com número alto).",
    )
    observacao = models.CharField("observação", max_length=200, blank=True)
    ativa = models.BooleanField(default=True)

    history = HistoricalRecords()

    class Meta:
        ordering = ["ordem", "id"]
        verbose_name = "regra"

    def __str__(self):
        return f"{self.palavra} → {self.conta}"


class Lancamento(models.Model):
    class Classificacao(models.TextChoices):
        AUTO = "A", "Automática"
        MANUAL = "M", "Manual"
        NENHUMA = "N", "Sem regra"

    ano = models.PositiveSmallIntegerField(db_index=True)
    mes = models.PositiveSmallIntegerField("mês", choices=MESES, db_index=True)
    secao = models.CharField("seção", max_length=1, choices=Secao.choices, default=Secao.DESPESA)
    descricao = models.CharField("descrição", max_length=255)
    valor = models.DecimalField(max_digits=14, decimal_places=2, null=True, blank=True)
    conta = models.ForeignKey(Conta, on_delete=models.PROTECT, null=True, blank=True, related_name="lancamentos",
                              verbose_name="plano de contas")
    classificacao = models.CharField(
        "classificação", max_length=1, choices=Classificacao.choices, default=Classificacao.NENHUMA,
    )
    regra = models.ForeignKey(
        Regra, on_delete=models.SET_NULL, null=True, blank=True, related_name="+",
        help_text="Regra que classificou automaticamente.",
    )
    linha_planilha = models.PositiveIntegerField("linha na planilha", null=True, blank=True)
    criado_em = models.DateTimeField(auto_now_add=True)
    atualizado_em = models.DateTimeField(auto_now=True)

    history = HistoricalRecords(excluded_fields=["atualizado_em"])

    class Meta:
        ordering = ["ano", "mes", "secao", "id"]
        verbose_name = "lançamento"

    def __str__(self):
        return f"{self.get_mes_display()}/{self.ano} – {self.descricao}"

    def aplicar_regras(self, regras=None):
        """Classifica pelas regras, a não ser que tenha sido corrigido à mão.

        Retorna True se a conta ou o status mudou.
        """
        if self.classificacao == self.Classificacao.MANUAL:
            return False
        from .classificador import classificar

        regra = classificar(self.descricao, self.secao, regras)
        antes = (self.conta_id, self.regra_id, self.classificacao)
        if regra:
            self.conta_id, self.regra_id, self.classificacao = regra.conta_id, regra.id, self.Classificacao.AUTO
        else:
            self.conta_id, self.regra_id, self.classificacao = None, None, self.Classificacao.NENHUMA
        return antes != (self.conta_id, self.regra_id, self.classificacao)
