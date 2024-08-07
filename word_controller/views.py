import os
import win32com.client
import pythoncom
from django.http import JsonResponse
from django.views.decorators.http import require_GET
from django.conf import settings

def initialize_com():
    """Initialize COM library"""
    pythoncom.CoInitialize()

def cleanup_com():
    """Cleanup COM library"""
    pythoncom.CoUninitialize()

# Global Word object and document
word_app = None
document = None
current_document = None

def get_word_instance():
    global word_app
    if not word_app:
        try:
            initialize_com()
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = True
        except Exception as e:
            return None, str(e)
    return word_app, None

@require_GET
def list_documents(request):
    doc_files = []
    for root, _, files in os.walk(settings.ROOT_WORD_DIR):
        for file in files:
            if file.endswith('.docx') or file.endswith('.doc'):
                doc_files.append(os.path.join(root, file))
    return JsonResponse({'documents': doc_files})

@require_GET
def open_document(request, file_name):
    global document, current_document
    app, error = get_word_instance()
    if error:
        return JsonResponse({"message": f"Failed to access Word: {error}"}, status=500)

    full_path = os.path.join(settings.ROOT_WORD_DIR, file_name)

    try:
        if document:
            document.Close(False)
        document = app.Documents.Open(full_path)
        app.ActiveWindow.Activate()  # Ensure the document is in the ActiveWindow state
        current_document = full_path
        return JsonResponse({"message": "Document loaded successfully.", "file_name": file_name})
    except Exception as e:
        return JsonResponse({"message": f"Failed to load document: {e}"}, status=500)

@require_GET
def close_document(request):
    global document, word_app
    if not document:
        return JsonResponse({"message": "No document is currently open"}, status=400)
    try:
        document.Close(False)  # Close the document without saving
        document = None
        current_document = None

        # If there are no more documents open, quit Word application
        if not word_app.Documents.Count:
            word_app.Quit()
            word_app = None
            cleanup_com()

        return JsonResponse({"message": "Document closed successfully."})
    except Exception as e:
        return JsonResponse({"message": f"Error closing document: {e}"}, status=500)

@require_GET
def scroll_up(request):
    global document
    if not document:
        return JsonResponse({"message": "No document loaded"}, status=400)
    try:
        word_app.ActiveWindow.SmallScroll(Down=-1)  # Scroll up
        return JsonResponse({"message": "Scrolled up."})
    except Exception as e:
        return JsonResponse({"message": f"Error scrolling up: {e}"}, status=500)

@require_GET
def scroll_down(request):
    global document
    if not document:
        return JsonResponse({"message": "No document loaded"}, status=400)
    try:
        word_app.ActiveWindow.SmallScroll(Down=1)  # Scroll down
        return JsonResponse({"message": "Scrolled down."})
    except Exception as e:
        return JsonResponse({"message": f"Error scrolling down: {e}"}, status=500)

@require_GET
def zoom_in(request):
    global document
    if not document:
        return JsonResponse({"message": "No document loaded"}, status=400)
    try:
        word_app.ActiveWindow.View.Zoom.Percentage += 10
        return JsonResponse({"message": "Zoomed in."})
    except Exception as e:
        return JsonResponse({"message": f"Error zooming in: {e}"}, status=500)

@require_GET
def zoom_out(request):
    global document
    if not document:
        return JsonResponse({"message": "No document loaded"}, status=400)
    try:
        word_app.ActiveWindow.View.Zoom.Percentage -= 10
        return JsonResponse({"message": "Zoomed out."})
    except Exception as e:
        return JsonResponse({"message": f"Error zooming out: {e}"}, status=500)

@require_GET
def enable_read_mode(request):
    global document
    if not document:
        return JsonResponse({"message": "No document loaded"}, status=400)
    try:
        word_app.ActiveWindow.View.ReadingLayout = True
        return JsonResponse({"message": "Read mode enabled."})
    except Exception as e:
        return JsonResponse({"message": f"Error enabling read mode: {e}"}, status=500)

@require_GET
def disable_read_mode(request):
    global document
    if not document:
        return JsonResponse({"message": "No document loaded"}, status=400)
    try:
        word_app.ActiveWindow.View.ReadingLayout = False
        return JsonResponse({"message": "Read mode disabled."})
    except Exception as e:
        return JsonResponse({"message": f"Error disabling read mode: {e}"}, status=500)
