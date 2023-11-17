from django.shortcuts import render
from django.http import HttpResponse

# Create your views here.
def index(request):
	return render(request, 'index.html')

def especialidadesyservicios(request):
	return render(request, 'especialidadesyservicios.html')

def sucursales(request):
	return render(request, 'sucursales.html')

def nosotros(request):
	return render(request, 'nosotros.html')

def contacto(request):
	return render(request, 'contacto.html')