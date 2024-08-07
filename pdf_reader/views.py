import os
import subprocess
import signal
import psutil
import pyautogui
from django.http import JsonResponse
from django.views.decorators.http import require_GET
from django.conf import settings

# Global Adobe Acrobat instance
current_pdf = None

ADOBE_ACROBAT_PATH = r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"

@require_GET
def list_pdfs(request):
    pdf_files = []
    for root, _, files in os.walk(settings.ROOT_PDF_DIR):
        for file in files:
            if file.endswith('.pdf'):
                pdf_files.append(os.path.join(root, file))
    return JsonResponse({'pdfs': pdf_files})

@require_GET
def open_pdf(request, file_name):
    global current_pdf
    full_path = os.path.join(settings.ROOT_PDF_DIR, file_name)

    # Debug: Print the full path to ensure it's correct
    print(f"Attempting to open PDF at path: {full_path}")

    try:
        if current_pdf:
            close_pdf(request)

        # Using a correctly formatted path for subprocess
        current_pdf = subprocess.Popen([ADOBE_ACROBAT_PATH, os.path.abspath(full_path)], shell=True)
        return JsonResponse({"message": "PDF loaded successfully.", "file_name": file_name})
    except Exception as e:
        return JsonResponse({"message": f"Failed to load PDF: {e}"}, status=500)

@require_GET
def scroll_up(request):
    try:
        pyautogui.scroll(200)  # Scroll up
        return JsonResponse({"message": "Scrolled up."})
    except Exception as e:
        return JsonResponse({"message": f"Error scrolling up: {e}"}, status=500)

@require_GET
def scroll_down(request):
    try:
        pyautogui.scroll(-200) # Scroll down
        return JsonResponse({"message": "Scrolled down."})
    except Exception as e:
        return JsonResponse({"message": f"Error scrolling down: {e}"}, status=500)

@require_GET
def zoom_in(request):
    try:
        pyautogui.keyDown('ctrl')
        pyautogui.press('=')
        pyautogui.keyUp('ctrl')
        return JsonResponse({"message": "Zoomed in."})
    except Exception as e:
        return JsonResponse({"message": f"Error zooming in: {e}"}, status=500)
@require_GET
def zoom_out(request):
    try:
        pyautogui.hotkey('ctrl', '-')  # Zoom out
        return JsonResponse({"message": "Zoomed out."})
    except Exception as e:
        return JsonResponse({"message": f"Error zooming out: {e}"}, status=500)

@require_GET
def enable_read_mode(request):
    try:
        pyautogui.hotkey('ctrl', 'h')  # Enable read mode (shortcut might differ)
        return JsonResponse({"message": "Read mode enabled."})
    except Exception as e:
        return JsonResponse({"message": f"Error enabling read mode: {e}"}, status=500)

@require_GET
def disable_read_mode(request):
    try:
        pyautogui.hotkey('ctrl', 'h')  # Disable read mode (shortcut might differ)
        return JsonResponse({"message": "Read mode disabled."})
    except Exception as e:
        return JsonResponse({"message": f"Error disabling read mode: {e}"}, status=500)

@require_GET
def save_pdf(request):
    try:
        pyautogui.hotkey('ctrl', 's')  # Save file
        return JsonResponse({"message": "PDF saved."})
    except Exception as e:
        return JsonResponse({"message": f"Error saving PDF: {e}"}, status=500)

@require_GET
def print_pdf(request):
    try:
        pyautogui.hotkey('ctrl', 'p')  # Print file
        return JsonResponse({"message": "Print dialog opened."})
    except Exception as e:
        return JsonResponse({"message": f"Error opening print dialog: {e}"}, status=500)


@require_GET
def close_pdf(request):
    global current_pdf
    if not current_pdf:
        return JsonResponse({"message": "No PDF loaded"}, status=400)

    try:
        # Attempt to close the Adobe Acrobat application using subprocess
        subprocess.run(["taskkill", "/IM", "Acrobat.exe", "/F"], check=True)

        # Check if the process is still running and force kill if necessary
        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['name'] == "Acrobat.exe":
                proc.kill()

        current_pdf = None
        return JsonResponse({"message": "PDF closed."})
    except Exception as e:
        return JsonResponse({"message": f"Error closing PDF: {e}"}, status=500)