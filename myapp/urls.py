from django.contrib import admin
from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name="index"),
    path('/especialidadesyservicios', views.especialidadesyservicios, name="especialidadesyservicios"),
    path('/sucursales', views.sucursales, name="sucursales"),
    path('/nosotros', views.nosotros, name="nosotros"),
    path('/contacto', views.contacto, name="contacto"),
]