from django.urls import path

from . import views

app_name = 'repasses'

urlpatterns = [
    path('', views.home, name='home'),
    path('medicos/', views.medicos, name='medicos'),
    path('importar/', views.importar, name='importar'),
    path('importar/exportar/', views.exportar, name='exportar'),
    path('saidas/<str:pasta>/<str:arquivo>', views.baixar_saida, name='baixar_saida'),
]
