import logging
import os
import time
import traceback

import win32com.client
import pythoncom
from django.http import JsonResponse
from django.views.decorators.http import require_GET, require_http_methods
from django.conf import settings

# Initialize logger
logger = logging.getLogger(__name__)

# Global PowerPoint object and presentation
powerpoint_app = None
presentation = None
current_presentation = None


def initialize_com():
    """Initialize COM library"""
    pythoncom.CoInitialize()

def cleanup_com():
    """Cleanup COM library"""
    pythoncom.CoUninitialize()

def cleanup_powerpoint():
    """Cleanup PowerPoint objects"""
    global powerpoint_app, presentation, current_presentation
    try:
        if presentation:
            presentation.Close()
            presentation = None
        if powerpoint_app:
            powerpoint_app.Quit()
            powerpoint_app = None
    except Exception as e:
        logger.error(f"Error cleaning up PowerPoint: {e}")
    finally:
        cleanup_com()



@require_GET
def powerpoint_controls(request):
    return JsonResponse({"message": "PowerPoint controls loaded", "status": "success"})


def get_powerpoint_instance():
    global powerpoint_app
    if not powerpoint_app:
        try:
            initialize_com()
            powerpoint_app = win32com.client.Dispatch("PowerPoint.Application")
            powerpoint_app.Visible = True
        except Exception as e:
            cleanup_powerpoint()  # Ensure cleanup on error
            return None, str(e)
    return powerpoint_app, None




def get_current_presentation(request):
    global current_presentation
    if current_presentation:
        return JsonResponse({"current_presentation": current_presentation})
    else:
        return JsonResponse({"current_presentation": None})


@require_http_methods(["GET"])
def open_presentation(request):
    global presentation, current_presentation
    app, error = get_powerpoint_instance()
    if error:
        return JsonResponse({"message": f"Failed to access PowerPoint: {error}"}, status=500)

    file_name = request.GET.get('file_name')
    if not file_name:
        return JsonResponse({"message": "File name not provided."}, status=400)

    full_path = os.path.join(settings.ROOT_DIR, file_name)
    full_path = os.path.normpath(full_path)

    if not os.path.exists(full_path):
        return JsonResponse({"message": f"File does not exist at path: {full_path}"}, status=400)

    try:
        if presentation:
            logger.info("Closing existing presentation.")
            presentation.Close()
        logger.info(f"Opening presentation at path: {full_path}")
        presentation = app.Presentations.Open(full_path)
        current_presentation = full_path
        return JsonResponse({"message": "Presentation loaded successfully.", "file_name": file_name, "show_powerpoint": True})
    except Exception as e:
        cleanup_powerpoint()  # Cleanup on error
        return JsonResponse({"message": f"Failed to load presentation: {e}"}, status=500)

@require_GET
def next_slide(request):
    global presentation
    initialize_com()  # Initialize COM
    try:
        if not presentation:
            return JsonResponse({"message": "No presentation loaded"}, status=400)

        slide_show_window = presentation.SlideShowWindow
        if not slide_show_window:
            logging.error("Slideshow window not found. Make sure the presentation is in slideshow mode.")
            return JsonResponse({"message": "Presentation is not in slideshow mode or slideshow window not found."}, status=400)

        slide_show_window.View.Next()
        return JsonResponse({"message": "Moved to next slide."})
    except AttributeError as e:
        logging.error(f"AttributeError moving to the next slide: {e}")
        return JsonResponse({"message": "Error: Presentation not in slideshow mode or PowerPoint application issue."}, status=500)
    except Exception as e:
        logging.error(f"Error moving to the next slide: {e}")
        return JsonResponse({"message": f"Error moving to the next slide: {str(e)}"}, status=500)
    finally:
        cleanup_com()  # Clean up COM

@require_GET
def prev_slide(request):
    global presentation
    initialize_com()  # Initialize COM
    try:
        if not presentation:
            return JsonResponse({"message": "No presentation loaded"}, status=400)

        slide_show_window = presentation.SlideShowWindow
        if not slide_show_window:
            logging.error("Slideshow window not found. Make sure the presentation is in slideshow mode.")
            return JsonResponse({"message": "Presentation is not in slideshow mode or slideshow window not found."}, status=400)

        slide_show_window.View.Previous()
        return JsonResponse({"message": "Moved to previous slide."})
    except AttributeError as e:
        logging.error(f"AttributeError moving to the previous slide: {e}")
        return JsonResponse({"message": "Error: Presentation not in slideshow mode or PowerPoint application issue."}, status=500)
    except Exception as e:
        logging.error(f"Error moving to the previous slide: {e}")
        return JsonResponse({"message": f"Error moving to the previous slide: {str(e)}"}, status=500)
    finally:
        cleanup_com()  # Clean up COM


@require_GET
def start_presentation(request):
    global presentation
    initialize_com();
    if not presentation:
        return JsonResponse({"message": "No presentation loaded"}, status=400)
    try:
        # Start the slideshow mode
        presentation.SlideShowSettings.Run()

        # Ensure the slideshow mode started successfully
        if not presentation.SlideShowWindow:
            return JsonResponse({"message": "Failed to start slideshow mode. Please check PowerPoint."}, status=500)

        return JsonResponse({"message": "Presentation started."})
    except Exception as e:
        logging.error(f"Error starting presentation: {e}")
        return JsonResponse({"message": f"Error starting presentation: {e}"}, status=500)
    finally:
        cleanup_com();

@require_GET
def end_presentation(request):
    global presentation, current_presentation
    initialize_com();
    if not presentation or not presentation.SlideShowWindow:
        return JsonResponse({"message": "No presentation loaded or presentation not in slideshow mode"}, status=400)
    try:
        presentation.SlideShowWindow.View.Exit()
        current_presentation = None
        return JsonResponse({"message": "Presentation ended."})
    except Exception as e:
        return JsonResponse({"message": f"Error ending presentation: {e}"}, status=500)
    finally:
        cleanup_com();
        if presentation:
            presentation.Close()
            presentation = None