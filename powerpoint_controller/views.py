import subprocess
from django.http import JsonResponse
from django.views.decorators.http import require_GET


def open_powerpoint(request):
    try:
        powerpoint_path = r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"  # Adjust this path if needed
        subprocess.run([powerpoint_path], check=True)
        return JsonResponse({"message": "Microsoft PowerPoint opened successfully."})
    except Exception as e:
        return JsonResponse({"message": f"Failed to open Microsoft PowerPoint: {e}"}, status=500)
