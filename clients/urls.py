from django.urls import path
from .views import loadClients

urlpatterns = [
    path("load_clients/", loadClients, name="load_clients")
]