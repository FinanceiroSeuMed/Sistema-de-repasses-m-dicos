from django.urls import path

from . import views

app_name = 'repasses'

urlpatterns = [
    path('', views.home, name='home'),
]
