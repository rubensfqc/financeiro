from django.contrib import admin
from simple_history.admin import SimpleHistoryAdmin

from .models import Banco, Conta, Lancamento, Regra

admin.site.site_header = "Gestão Financeira – administração"
admin.site.site_title = "Gestão Financeira"


@admin.register(Conta)
class ContaAdmin(SimpleHistoryAdmin):
    list_display = ("codigo", "nome", "nao_entra_no_dre", "ativa")
    list_filter = ("ativa", "nao_entra_no_dre")
    search_fields = ("codigo", "nome")


@admin.register(Banco)
class BancoAdmin(SimpleHistoryAdmin):
    list_display = ("nome", "coluna_planilha", "banco", "agencia", "numero_conta", "ativo")
    search_fields = ("nome", "numero_conta")


@admin.register(Regra)
class RegraAdmin(SimpleHistoryAdmin):
    list_display = ("ordem", "palavra", "tipo", "secao", "conta", "ativa")
    list_filter = ("tipo", "secao", "ativa", "conta")
    search_fields = ("palavra", "observacao")
    autocomplete_fields = ("conta",)

    # Mudou uma regra → reaplica a todos os lançamentos automáticos.
    def save_model(self, request, obj, form, change):
        super().save_model(request, obj, form, change)
        from .classificador import reclassificar
        reclassificar(motivo=f"Regra alterada: {obj.palavra}", usuario=request.user)

    def delete_model(self, request, obj):
        nome = obj.palavra
        super().delete_model(request, obj)
        from .classificador import reclassificar
        reclassificar(motivo=f"Regra excluída: {nome}", usuario=request.user)


@admin.register(Lancamento)
class LancamentoAdmin(SimpleHistoryAdmin):
    list_display = ("ano", "mes", "secao", "descricao", "valor", "banco", "conta", "classificacao")
    list_filter = ("ano", "mes", "secao", "classificacao", "banco", "conta")
    search_fields = ("descricao",)
    autocomplete_fields = ("conta",)
    readonly_fields = ("regra", "criado_em", "atualizado_em")
