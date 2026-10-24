from django import forms
from django.utils import timezone

from .models import MESES, Conta, Lancamento, Regra


class LancamentoForm(forms.ModelForm):
    conta = forms.ModelChoiceField(
        queryset=Conta.objects.filter(ativa=True), required=False,
        empty_label="— automática (pelas regras) —",
        help_text="Deixe em automática para a conta vir das regras. Escolher uma conta trava a classificação.",
    )

    class Meta:
        model = Lancamento
        fields = ["ano", "mes", "secao", "descricao", "valor", "conta"]
        widgets = {"descricao": forms.TextInput(attrs={"autofocus": True, "autocomplete": "off"})}

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        if not self.instance.pk:
            hoje = timezone.localdate()
            self.fields["ano"].initial = hoje.year
            self.fields["mes"].initial = hoje.month
        elif self.instance.classificacao != Lancamento.Classificacao.MANUAL:
            self.initial["conta"] = None

    def save(self, commit=True):
        lanc = super().save(commit=False)
        if self.cleaned_data.get("conta"):
            lanc.classificacao = Lancamento.Classificacao.MANUAL
            lanc.regra = None
        else:
            lanc.classificacao = Lancamento.Classificacao.NENHUMA  # deixa as regras decidirem
            lanc.aplicar_regras()
        if commit:
            lanc.save()
        return lanc


class RegraForm(forms.ModelForm):
    class Meta:
        model = Regra
        fields = ["palavra", "tipo", "secao", "conta", "ordem", "observacao", "ativa"]

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields["conta"].queryset = Conta.objects.filter(ativa=True)
        if not self.instance.pk:
            ultima = Regra.objects.order_by("-ordem").values_list("ordem", flat=True).first() or 0
            self.fields["ordem"].initial = ultima + 1

    def clean_palavra(self):
        return self.cleaned_data["palavra"].strip().upper()


class ContaForm(forms.ModelForm):
    class Meta:
        model = Conta
        fields = ["codigo", "nome", "nao_entra_no_dre", "ativa"]


class ImportarForm(forms.Form):
    arquivo = forms.FileField(label="Planilha (.xlsx)")
    senha = forms.CharField(label="Senha da planilha", required=False, widget=forms.PasswordInput(render_value=False))
    ano = forms.IntegerField(min_value=2000, max_value=2100)
    importar_regras = forms.BooleanField(
        label="Importar regras e contas (substitui as regras atuais)", required=False, initial=True)
    importar_lancamentos = forms.BooleanField(
        label="Importar lançamentos (substitui os lançamentos deste ano)", required=False, initial=True)
    coluna_valor = forms.CharField(
        label="Coluna do valor (opcional)", required=False, max_length=3,
        help_text="Letra da coluna com o valor de cada lançamento nas abas de mês (ex.: D). Em branco = não importar valores.",
    )

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields["ano"].initial = timezone.localdate().year

    def clean_coluna_valor(self):
        v = self.cleaned_data["coluna_valor"].strip().upper()
        if v and not v.isalpha():
            raise forms.ValidationError("Use só a letra da coluna, ex.: D.")
        return v


MESES_FILTRO = [("", "Todos")] + [(str(n), nome) for n, nome in MESES]
