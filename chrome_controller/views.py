# chrome_controller/views.py
import os

import win32com.client
import pythoncom
from django.http import JsonResponse
from django.views.decorators.http import require_GET

def initialize_com():
    """Initialize COM library"""
    pythoncom.CoInitialize()

def cleanup_com():
    """Cleanup COM library"""
    pythoncom.CoUninitialize()

# Global Chrome object
chrome_app = None

def get_chrome_instance():
    global chrome_app
    if not chrome_app:
        try:
            initialize_com()
            chrome_app = win32com.client.Dispatch("Chrome.Application")
            chrome_app.Visible = True
        except Exception as e:
            return None, str(e)
    return chrome_app, None


@require_GET
def open_chrome(request):
    try:
        # Path to Chrome executable (modify if needed)
        chrome_path = 'C:/Program Files/Google/Chrome/Application/chrome.exe'

        # Command to open Chrome
        os.startfile(chrome_path)
        return JsonResponse({"message": "Chrome opened successfully."})
    except Exception as e:
        return JsonResponse({"message": f"Error opening Chrome: {e}"}, status=500)

@require_GET
def go_home(request):
    app, error = get_chrome_instance()
    if error:
        return JsonResponse({"message": f"Failed to access Chrome: {error}"}, status=500)
    try:
        app.ActiveWindow.Navigate("about:home")
        return JsonResponse({"message": "Navigated to home."})
    except Exception as e:
        return JsonResponse({"message": f"Failed to navigate home: {e}"}, status=500)

@require_GET
def go_back(request):
    app, error = get_chrome_instance()
    if error:
        return JsonResponse({"message": f"Failed to access Chrome: {error}"}, status=500)
    try:
        app.ActiveWindow.GoBack()
        return JsonResponse({"message": "Navigated back."})
    except Exception as e:
        return JsonResponse({"message": f"Failed to navigate back: {e}"}, status=500)

@require_GET
def go_forward(request):
    app, error = get_chrome_instance()
    if error:
        return JsonResponse({"message": f"Failed to access Chrome: {error}"}, status=500)
    try:
        app.ActiveWindow.GoForward()
        return JsonResponse({"message": "Navigated forward."})
    except Exception as e:
        return JsonResponse({"message": f"Failed to navigate forward: {e}"}, status=500)

@require_GET
def close_chrome(request):
    global chrome_app
    if not chrome_app:
        return JsonResponse({"message": "Chrome is not running"}, status=400)
    try:
        chrome_app.Quit()
        chrome_app = None
        cleanup_com()
        return JsonResponse({"message": "Chrome closed successfully."})
    except Exception as e:
        return JsonResponse({"message": f"Failed to close Chrome: {e}"}, status=500)

# Add similar views for other actions like zooming, scrolling, and tab management
