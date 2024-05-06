import os
import subprocess
import sys
from django.conf import settings
from django.shortcuts import render
from django.http import HttpResponse, JsonResponse
from django.core.files.storage import FileSystemStorage
import shutil
import glob

def delete_files_in_directory(directory):
    files = glob.glob(os.path.join(directory, '*'))
    for file in files:
        os.remove(file)

def upload_and_process_documents(request):
    try:
        if request.method == 'POST':
            use_pip_checkbox = request.POST.get('use_pip_yesterday', '') == 'on'
            use_vector_checkbox = request.POST.get('use_vector_yesterday', '') == 'on'

            pip_file = os.path.join(settings.DATA_A_EXTRAER_DIR, 'PIP.xls')
            pip_renamed_file = os.path.join(settings.DATA_A_EXTRAER_DIR, 'PipViejo.xls')
            vector_file = os.path.join(settings.DATA_A_EXTRAER_DIR, 'VectorAnalitico24h.xls')
            vector_renamed_file = os.path.join(settings.DATA_A_EXTRAER_DIR, 'VectorViejo.xls')

            if use_pip_checkbox and os.path.exists(pip_file):
                shutil.move(pip_file, pip_renamed_file)

            if use_vector_checkbox and os.path.exists(vector_file):
                shutil.move(vector_file, vector_renamed_file)

            files = glob.glob(os.path.join(settings.DATA_A_EXTRAER_DIR, '*'))
            for file in files:
                if file not in [pip_renamed_file, vector_renamed_file]:
                    os.remove(file)

            delete_files_in_directory(settings.OUTPUT_DIR)
            delete_files_in_directory(settings.EXCELES_DIR)

            files = request.FILES.getlist('documents')
            fs = FileSystemStorage(location=settings.DATA_A_EXTRAER_DIR)

            # Save each file in the specified directory
            for file in files:
                fs.save(file.name, file)

            # Execute the Python script using the interpreter of the current environment
            script_path = os.path.join(settings.FUNCIONALIDAD_DIR, 'Programas', 'main_df.py')
            subprocess.run([sys.executable, script_path], check=True)

            # Collect output files for download
            output_files = os.listdir(settings.OUTPUT_DIR)
            return render(request, 'download_output.html', {'files': output_files})

    except FileNotFoundError as fnf_error:
        return JsonResponse({'status': 'error', 'message': f"No se encontró un archivo: {fnf_error}"})

    except subprocess.CalledProcessError as cpe_error:
        return JsonResponse({'status': 'error', 'message': "Hubo un error al correr el programa. Es probable que no hayas subido los archivos correctamente. Recuerda que los nombres de los documentos subidos deben ser exactamente iguales a los que se detallan en la lista de documentos."})

    except Exception as generic_error:
        return JsonResponse({'status': 'error', 'message': f"Ocurrió un error inesperado: {generic_error}"})

    return render(request, 'upload_documents.html')

def download_file(request, filename):
    try:
        file_path = os.path.join(settings.OUTPUT_DIR, filename)
        response = HttpResponse(open(file_path, 'rb'), content_type='application/vnd.ms-excel')
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        return response
    except FileNotFoundError as fnf_error:
        return HttpResponse(f"No se encontró el archivo final: {fnf_error}")
    except Exception as generic_error:
        return HttpResponse(f"Ocurrió un error inesperado: {generic_error}")