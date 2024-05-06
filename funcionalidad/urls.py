from . import views 
from django.urls import path


urlpatterns = [
    path('', views.input_documentos, name='input_documentos'),
]