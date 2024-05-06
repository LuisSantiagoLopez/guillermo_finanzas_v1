import os
import subprocess
import sys
from django.conf import settings
from django.shortcuts import render
from django.http import HttpResponse
from django.core.files.storage import FileSystemStorage
import glob

def delete_files_in_directory(directory):
    files = glob.glob(os.path.join(directory, '*'))
    for file in files:
        os.remove(file)

def upload_and_process_documents(request):
    if request.method == 'POST': 
        delete_files_in_directory(settings.DATA_A_EXTRAER_DIR)
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

    return render(request, 'upload_documents.html')

def download_file(request, filename):
    file_path = os.path.join(settings.OUTPUT_DIR, filename)
    response = HttpResponse(open(file_path, 'rb'), content_type='application/vnd.ms-excel')
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    return response
