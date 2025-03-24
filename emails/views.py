from django.shortcuts import render
from clients.models import Client
from django.http import HttpResponse
from django.views.decorators.csrf import csrf_exempt
from .utils import email_for_clients

@csrf_exempt
def send_emails(request):
    clientes = Client.objects.all()
    print(clientes)
    for client in clientes:
        estado = email_for_clients(client)

    return HttpResponse("enviados correctamente")