from django.urls import path

from . import views

urlpatterns = [
    path("", views.lancamentos, name="lancamentos"),
    path("lancamentos/novo/", views.lancamento_form, name="lancamento_novo"),
    path("lancamentos/<int:pk>/", views.lancamento_form, name="lancamento_editar"),
    path("lancamentos/<int:pk>/excluir/", views.lancamento_excluir, name="lancamento_excluir"),
    path("lancamentos/<int:pk>/conta/", views.lancamento_conta, name="lancamento_conta"),
    path("sugerir/", views.sugerir, name="sugerir"),
    path("reclassificar/", views.reclassificar_tudo, name="reclassificar"),
    path("regras/", views.regras, name="regras"),
    path("regras/nova/", views.regra_form, name="regra_nova"),
    path("regras/<int:pk>/", views.regra_form, name="regra_editar"),
    path("regras/<int:pk>/excluir/", views.regra_excluir, name="regra_excluir"),
    path("contas/", views.contas, name="contas"),
    path("contas/nova/", views.conta_form, name="conta_nova"),
    path("contas/<int:pk>/", views.conta_form, name="conta_editar"),
    path("resumo/", views.resumo, name="resumo"),
    path("historico/", views.historico, name="historico"),
    path("importar/", views.importar, name="importar"),
    path("exportar/", views.exportar, name="exportar"),
]
