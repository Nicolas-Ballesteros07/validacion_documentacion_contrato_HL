from django.urls import path

from .views import validar_documentos

urlpatterns = [
    path(
        '',
        validar_documentos,
        name='validar_documentos'
    ),
]