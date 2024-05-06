from django.urls import path
from . import views

urlpatterns = [
    path('upload/', views.upload_and_process_documents, name='upload_and_process_documents'),
    path('download/<str:filename>/', views.download_file, name='download_file'),
]
