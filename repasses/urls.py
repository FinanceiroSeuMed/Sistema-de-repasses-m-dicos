from django.urls import path

from . import views

app_name = 'repasses'

urlpatterns = [
    path('', views.home, name='home'),
    path('medicos/', views.medicos, name='medicos'),
    path('importar/', views.importar, name='importar'),
    path('importar/salvar/', views.salvar, name='salvar'),
    path('importar/visualizar/', views.visualizar, name='visualizar'),
    path('importar/revisar/', views.revisar, name='revisar'),
    path('importar/cadastrar-medicos/', views.cadastrar_medicos, name='cadastrar_medicos'),
    path('importar/continuar/', views.continuar_edicao, name='continuar_edicao'),
    path('importar/exportar/', views.exportar, name='exportar'),
    path('saidas/<str:pasta>/<str:arquivo>', views.baixar_saida, name='baixar_saida'),
    path('lotes/', views.lotes_lista, name='lotes'),
    path('lotes/sem-repasse/', views.confirmar_sem_repasse, name='confirmar_sem_repasse'),
    path('lotes/sem-repasse/<int:pk>/remover/', views.remover_sem_repasse, name='remover_sem_repasse'),
    path('lotes/relatorio-dias/', views.relatorio_dias, name='relatorio_dias'),
    path('relatorio-mensal/', views.relatorio_mensal, name='relatorio_mensal'),
    path('relatorio-mensal/ajustes/', views.salvar_ajuste_mensal, name='salvar_ajuste_mensal'),
    path('lotes/<int:pk>/', views.lote_detalhe, name='lote_detalhe'),
    path('lotes/<int:pk>/status/', views.lote_status, name='lote_status'),
    path('lotes/<int:pk>/reabrir/', views.lote_reabrir, name='lote_reabrir'),
    path('lotes/<int:pk>/excluir/', views.lote_excluir, name='lote_excluir'),
    path('lotes/<int:pk>/baixar/<str:nome>', views.baixar_lote, name='baixar_lote'),
    path('lotes/<int:pk>/baixar-zip/', views.baixar_lote_zip, name='baixar_lote_zip'),
    path('administracao/', views.administracao, name='administracao'),
    path('regras/', views.regras_lista, name='regras'),
    path('regras/salvar/', views.regras_salvar, name='regras_salvar'),
    path('correcoes/', views.correcoes_lista, name='correcoes'),
    path('correcoes/<int:pk>/ligar/', views.correcao_toggle, name='correcao_toggle'),
    path('correcoes/<int:pk>/remover/', views.correcao_remover, name='correcao_remover'),
    path('classes/<int:pk>/ligar/', views.classe_toggle, name='classe_toggle'),
    path('classes/<int:pk>/remover/', views.classe_remover, name='classe_remover'),
]
