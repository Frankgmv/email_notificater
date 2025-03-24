from django.db import models


class Client(models.Model):
    correo1 = models.CharField(max_length=100)
    correo2 = models.EmailField(blank=True, null=True)
    nit = models.CharField(max_length=100, unique=True)
    nombre_empresa = models.CharField(max_length=100)
    nombre_gerente = models.CharField(max_length=100)
    correo_gerente = models.CharField(max_length=100)
    hecho = models.BooleanField(default=False)
    

    def __str__(self):
        return self.nit + " - " + self.nombre_empresa