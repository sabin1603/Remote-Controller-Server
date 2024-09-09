import os
import subprocess

from django.http import JsonResponse, FileResponse
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_GET, require_POST
from django.shortcuts import render

@require_GET
def list_files(request):
    base_path = request.GET.get('path', '/')
    if not os.path.exists(base_path):
        return JsonResponse({'error': 'Path does not exist'}, status=400)

    file_tree = []
    for root, dirs, files in os.walk(base_path):
        for name in dirs + files:
            path = os.path.join(root, name)
            file_tree.append({
                'name': name,
                'path': path,
                'extension': os.path.splitext(name)[1] if os.path.isfile(path) else 'folder'
            })
        break

    return JsonResponse(file_tree, safe=False)


@require_POST
@csrf_exempt
def open_file(request):
    file_path = request.POST.get('path')
    if not file_path or not os.path.exists(file_path):
        return JsonResponse({'error': 'File does not exist'}, status=400)

    try:
        # Determine the module based on the file extension
        extension = os.path.splitext(file_path)[1].lower()
        module = None

        if extension == '.pptx':
            module = 'powerpoint'
        elif extension == '.docx':
            module = 'word'
        elif extension == '.xlsx':
            module = 'excel'
        elif extension == '.pdf':
            module = 'pdf-reader'

        # Open the file on the server-side
        if os.name == 'nt':  # Windows
            os.startfile(file_path)
        elif os.name == 'posix':  # macOS or Linux
            if os.uname().sysname == 'Darwin':  # macOS
                subprocess.call(['open', file_path])
            else:  # Linux
                subprocess.call(['xdg-open', file_path])

        # Respond with the appropriate module URL
        if module:
            return JsonResponse({'message': 'File opened successfully', 'module': module})
        else:
            return JsonResponse({'message': 'File opened successfully, but no module found for this file type.'})

    except Exception as e:
        return JsonResponse({'error': str(e)}, status=500)


@require_GET
def download_file(request):
    file_path = request.GET.get('path')
    if not file_path or not os.path.exists(file_path):
        return JsonResponse({'error': 'File does not exist'}, status=400)

    try:
        response = FileResponse(open(file_path, 'rb'), as_attachment=True)
        response['Content-Disposition'] = f'attachment; filename="{os.path.basename(file_path)}"'
        return response
    except Exception as e:
        return JsonResponse({'error': str(e)}, status=500)

def index(request):
    return render(request, 'index.html')
