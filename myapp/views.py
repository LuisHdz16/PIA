from django.shortcuts import render
from django.http import HttpResponse
from .models import Medico, Sucursal, Ambulancia, Departamento

# Create your views here.
def index(request):
	return render(request, 'index.html')

def especialidadesyservicios(request):
	departamento = Departamento.objects.all()
	ambulancia = Ambulancia.objects.all()
	return render(request, 'especialidadesyservicios.html', {
		"departamentos" : departamento,
		"ambulancias" : ambulancia
	})

def sucursales(request):
	sucursal = Sucursal.objects.all()
	return render(request, 'sucursales.html', {
		"sucursales" : sucursal,
	})

def nosotros(request):
	data = Medico.objects.all()
	return render(request, 'nosotros.html', {
		"medicos" : data
	})

def contacto(request):
	return render(request, 'contacto.html')