from django.contrib import admin
from .models import Departamento, Sucursal, Medico, Ambulancia

# Register your models here.
admin.site.register(Departamento)
admin.site.register(Sucursal)
admin.site.register(Medico)
admin.site.register(Ambulancia)