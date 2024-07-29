import os
import win32com.client
import pythoncom
from django.http import JsonResponse
from django.views.decorators.http import require_GET
from django.shortcuts import render
from .utils import find_pptx_files
from django.conf import settings

def initialize_com():
    """Initialize COM library"""
    pythoncom.CoInitialize()


def cleanup_com():
    """Cleanup COM library"""
    pythoncom.CoUninitialize()


# Global PowerPoint object and presentation
powerpoint_app = None
presentation = None
current_presentation = None


def get_powerpoint_instance():
    global powerpoint_app
    if not powerpoint_app:
        try:
            initialize_com()
            powerpoint_app = win32com.client.Dispatch("PowerPoint.Application")
            powerpoint_app.Visible = True
        except Exception as e:
            return None, str(e)
    return powerpoint_app, None



@require_GET
def list_presentations(request):
    pptx_files = []
    for root, dirs, files in os.walk(settings.ROOT_DIR):
        for file in files:
            if file.endswith('.pptx'):
                pptx_files.append(os.path.join(root, file))
    return JsonResponse({'presentations': pptx_files})

@require_GET
def open_presentation(request, file_name):
    """Open a selected presentation"""
    global presentation, current_presentation
    app, error = get_powerpoint_instance()
    if error:
        return JsonResponse({"message": f"Failed to access PowerPoint: {error}"}, status=500)

    full_path = file_name

    try:
        if presentation:
            presentation.Close()
        presentation = app.Presentations.Open(full_path)
        current_presentation = file_name
        return JsonResponse({"message": "Presentation loaded successfully.", "file_name": file_name})
    except Exception as e:
        return JsonResponse({"message": f"Failed to load presentation: {e}"}, status=500)


@require_GET
def next_slide(request):
    global presentation
    if not presentation or not presentation.SlideShowWindow:
        return JsonResponse({"message": "No presentation loaded or presentation not in slideshow mode"}, status=400)
    try:
        presentation.SlideShowWindow.View.Next()
        return JsonResponse({"message": "Moved to next slide."})
    except Exception as e:
        return JsonResponse({"message": f"Error moving to the next slide: {e}"}, status=500)


@require_GET
def prev_slide(request):
    global presentation
    if not presentation or not presentation.SlideShowWindow:
        return JsonResponse({"message": "No presentation loaded or presentation not in slideshow mode"}, status=400)
    try:
        presentation.SlideShowWindow.View.Previous()
        return JsonResponse({"message": "Moved to previous slide."})
    except Exception as e:
        return JsonResponse({"message": f"Error moving to the previous slide: {e}"}, status=500)


@require_GET
def start_presentation(request):
    global presentation
    if not presentation:
        return JsonResponse({"message": "No presentation loaded"}, status=400)
    try:
        presentation.SlideShowSettings.Run()
        return JsonResponse({"message": "Presentation started."})
    except Exception as e:
        return JsonResponse({"message": f"Error starting presentation: {e}"}, status=500)


@require_GET
def end_presentation(request):
    global presentation, current_presentation
    if not presentation or not presentation.SlideShowWindow:
        return JsonResponse({"message": "No presentation loaded or presentation not in slideshow mode"}, status=400)
    try:
        presentation.SlideShowWindow.View.Exit()
        current_presentation = None
        return JsonResponse({"message": "Presentation ended."})
    except Exception as e:
        return JsonResponse({"message": f"Error ending presentation: {e}"}, status=500)
    finally:
        if presentation:
            presentation.Close()
