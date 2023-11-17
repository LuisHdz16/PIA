from django.contrib import admin
from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name="index"),
    path('', views.especialidadesyservicios, name="especialidadesyservicios"),
    path('', views.sucursales, name="sucursales"),
    path('', views.nosotros, name="nosotros"),
    path('', views.contacto, name="contacto"),
]