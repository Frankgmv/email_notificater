from django.shortcuts import render
from django.views.decorators.csrf import csrf_exempt
from django.http import HttpResponse, JsonResponse
from .models import Client
import json


@csrf_exempt
def loadClients(request):
    try:

        data = json.loads(request.body)
        if not isinstance(data, list):
            return JsonResponse(
                {"error": "Se esperaba un array de clientes"}, status=400
            )

        for cliente in data:
            correo1 = cliente.get("correo1")
            correo2 = cliente.get("correo2")
            nombre_cliente = cliente.get("nombre_cliente")
            nit = cliente.get("nit")
            nombre_empresa = cliente.get("nombre_empresa")
            nombre_gerente = cliente.get("nombre_gerente")
            correo_gerente = cliente.get("correo_gerente")

            # Aquí puedes guardar en tu modelo de Django, por ejemplo:
            Client.objects.create(
                correo1=correo1,
                correo2=correo2,
                nombre_cliente=nombre_cliente,
                nit=nit,
                nombre_empresa=nombre_empresa,
                nombre_gerente=nombre_gerente,
                correo_gerente=correo_gerente
            )

            print(f"Client received: {nombre_cliente} - {nombre_empresa}")
        return HttpResponse("Clients created successfuly")
    except json.JSONDecodeError:
        return JsonResponse({"error": "Formato JSON inválido"}, status=400)
    except Exception as e:
        return JsonResponse({"error": str(e)}, status=500)