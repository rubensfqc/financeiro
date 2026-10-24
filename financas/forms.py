from django import forms
from django.utils import timezone

from .models import MESES, Conta, Lancamento, Regra


class LancamentoForm(forms.ModelForm):
    conta = forms.ModelChoiceField(
        label="Plano de contas",
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
    mesclar_em = forms.ModelChoiceField(
        label="Mesclar com outra conta", queryset=Conta.objects.none(), required=False,
        empty_label="— não mesclar —",
        help_text="Move todos os lançamentos e regras desta conta para a escolhida e apaga esta. "
                  "Use quando a planilha trouxe o mesmo item com outro nome.",
    )

    class Meta:
        model = Conta
        fields = ["codigo", "nome", "nao_entra_no_dre", "ativa"]

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        if self.instance.pk:
            self.fields["mesclar_em"].queryset = Conta.objects.exclude(pk=self.instance.pk).exclude(codigo="")
        else:
            del self.fields["mesclar_em"]

    def clean_codigo(self):
        codigo = self.cleaned_data["codigo"].strip()
        if codigo and Conta.objects.filter(codigo=codigo).exclude(pk=self.instance.pk).exists():
            raise forms.ValidationError("Já existe outra conta com este código. Para juntar as duas, use “Mesclar”.")
        return codigo


class ImportarForm(forms.Form):
    MODOS = [
        ("completa", "Importação completa: regras, plano de contas, lançamentos e valores "
                     "(substitui os lançamentos do ano)"),
        ("valores", "Só atualizar os valores dos lançamentos já importados "
                    "(mantém contas e correções feitas no app)"),
    ]
    arquivo = forms.FileField(label="Planilha (.xlsx)")
    senha = forms.CharField(label="Senha da planilha", required=False, widget=forms.PasswordInput(render_value=False))
    ano = forms.IntegerField(min_value=2000, max_value=2100)
    modo = forms.ChoiceField(label="O que fazer", choices=MODOS, initial="completa", widget=forms.RadioSelect,
                             required=False)
    importar_regras = forms.BooleanField(
        label="Na importação completa, importar também as regras (substitui as regras atuais)",
        required=False, initial=True)
    coluna_valor = forms.CharField(
        label="Coluna do valor", max_length=3, initial="D",
        help_text="Letra da coluna com o valor de cada lançamento nas abas de mês (na planilha atual: D).",
    )

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields["ano"].initial = timezone.localdate().year

    def clean_modo(self):
        return self.cleaned_data.get("modo") or "completa"

    def clean_coluna_valor(self):
        v = self.cleaned_data["coluna_valor"].strip().upper()
        if not v.isalpha():
            raise forms.ValidationError("Use só a letra da coluna, ex.: D.")
        return v


MESES_FILTRO = [("", "Todos")] + [(str(n), nome) for n, nome in MESES]
