from django.db import models

# Create your models here.
class Departamento(models.Model):
    nombre_departamento = models.CharField(max_length=50)
    jefe_departamento = models.CharField(max_length=100)
    ubicacion_departamento = models.CharField(max_length=100)
    telefono_departamento = models.CharField(max_length=20)
    numero_empleados = models.IntegerField()

    def __str__(self):
        return self.nombre_departamento
        return self.jefe_departamento
        return self.ubicacion_departamento
        return self.telefono_departamento
        return self.numero_empleados

class Sucursal(models.Model):
    nombre_sucursal = models.CharField(max_length=50)
    responsable_sucursal = models.CharField(max_length=100)
    ubicacion_sucursal = models.CharField(max_length=100)
    telefono_sucursal = models.CharField(max_length=20)
    horario_atencion = models.CharField(max_length=100)

    def __str__(self):
        return self.nombre_sucursal
        return self.responsable_sucursal
        return self.ubicacion_sucursal
        return self.telefono_sucursal
        return self.horario_atencion

class Medico(models.Model):
    nombre_medico = models.CharField(max_length=50)
    especialidad_medico = models.CharField(max_length=50)
    fecha_nacimiento = models.DateField()
    años_experiencia = models.IntegerField()
    correo_electronico = models.CharField(max_length=150)

    def __str__(self):
        return self.nombre_medico
        return self.especialidad_medico
        return self.fecha_nacimiento
        return self.años_experiencia
        return self.correo_electronico

class Ambulancia(models.Model):
    modelo = models.CharField(max_length=50)
    placa = models.CharField(max_length=20)
    fecha_compra = models.DateField()
    estado = models.CharField(max_length=20)
    cantidad = models.IntegerField()

    def __str__(self):
        return self.modelo
        return self.placa
        return self.fecha_compra
        return self.estado
        return self.cantidad