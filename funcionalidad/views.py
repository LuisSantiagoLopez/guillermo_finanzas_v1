from django.shortcuts import render


def input_documentos(request):
    return render(request, 'input_documentos.html')