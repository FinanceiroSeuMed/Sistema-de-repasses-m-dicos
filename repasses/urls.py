from django.urls import path

from . import views

app_name = 'repasses'

urlpatterns = [
    path('', views.home, name='home'),
    path('importar/', views.importar, name='importar'),
    path('saidas/<str:pasta>/<str:arquivo>', views.baixar_saida, name='baixar_saida'),
]
