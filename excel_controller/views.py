import logging
import os
import win32com.client
import pythoncom
from django.http import JsonResponse
from django.views.decorators.http import require_GET, require_http_methods
from django.conf import settings

# Initialize logger
logger = logging.getLogger(__name__)


# Initialize COM library
def initialize_com():
    pythoncom.CoInitialize()


# Cleanup COM library
def cleanup_com():
    pythoncom.CoUninitialize()


# Initialize global variables
excel_app = None
workbooks = {}
current_workbook = None  # New variable to keep track of the currently active workbook


def get_excel_instance():
    global excel_app
    if not excel_app:
        try:
            initialize_com()
            excel_app = win32com.client.Dispatch("Excel.Application")
            excel_app.Visible = True
        except Exception as e:
            logger.error(f"Failed to initialize Excel: {e}")
            return None, str(e)
    return excel_app, None


@require_http_methods(["GET"])
def open_workbook(request):
    global workbooks, current_workbook
    app, error = get_excel_instance()
    if error:
        return JsonResponse({"message": f"Failed to access Excel: {error}"}, status=500)

    file_name = request.GET.get('file_name')
    if not file_name:
        return JsonResponse({"message": "File name not provided."}, status=400)

    full_path = os.path.join(settings.ROOT_DIR, file_name)
    full_path = os.path.normpath(full_path)

    if not os.path.exists(full_path):
        return JsonResponse({"message": f"File does not exist at path: {full_path}"}, status=400)

    try:
        # Check if the workbook is already open
        for wb in app.Workbooks:
            if wb.FullName == full_path:
                workbooks[file_name] = wb
                current_workbook = wb  # Set the current workbook
                logger.info(f"Workbook already open: {file_name}")
                return JsonResponse(
                    {"message": "Workbook is already open.", "file_name": file_name, "show_excel": True})

        # Open the workbook if not already open
        workbook = app.Workbooks.Open(full_path)
        workbooks[file_name] = workbook
        current_workbook = workbook  # Set the current workbook
        logger.info(f"Opened workbook: {file_name}")
        return JsonResponse({"message": "Workbook loaded successfully.", "file_name": file_name, "show_excel": True})
    except Exception as e:
        cleanup_com()  # Cleanup on error
        return JsonResponse({"message": f"Failed to load workbook: {e}"}, status=500)


@require_GET
def close_workbook(request):
    global workbooks, excel_app, current_workbook
    file_name = request.GET.get('file_name')
    if not file_name or file_name not in workbooks:
        return JsonResponse({"message": "No workbook is currently open"}, status=400)

    try:
        workbook = workbooks[file_name]
        workbook.Close(False)  # Close the workbook without saving
        del workbooks[file_name]

        # Update the current workbook if it was closed
        if current_workbook == workbook:
            current_workbook = None

        logger.info(f"Closed workbook: {file_name}")

        # If there are no more workbooks open, quit Excel application
        if not workbooks:
            excel_app.Quit()
            excel_app = None
            cleanup_com()

        return JsonResponse({"message": "Workbook closed successfully."})
    except Exception as e:
        return JsonResponse({"message": f"Error closing workbook: {e}"}, status=500)


def get_current_workbook():
    if current_workbook:
        return current_workbook
    return None


@require_GET
def next_worksheet(request):
    global workbook
    if not workbook:
        return JsonResponse({"message": "No workbook loaded"}, status=400)
    try:
        current_sheet_index = workbook.ActiveSheet.Index
        next_sheet_index = current_sheet_index + 1
        if next_sheet_index <= workbook.Sheets.Count:
            workbook.Sheets(next_sheet_index).Activate()
            return JsonResponse({"message": "Moved to next worksheet."})
        else:
            return JsonResponse({"message": "Already on the last worksheet."}, status=400)
    except Exception as e:
        return JsonResponse({"message": f"Error moving to the next worksheet: {e}"}, status=500)

@require_GET
def previous_worksheet(request):
    global workbook
    if not workbook:
        return JsonResponse({"message": "No workbook loaded"}, status=400)
    try:
        current_sheet_index = workbook.ActiveSheet.Index
        previous_sheet_index = current_sheet_index - 1
        if previous_sheet_index >= 1:
            workbook.Sheets(previous_sheet_index).Activate()
            return JsonResponse({"message": "Moved to previous worksheet."})
        else:
            return JsonResponse({"message": "Already on the first worksheet."}, status=400)
    except Exception as e:
        return JsonResponse({"message": f"Error moving to the previous worksheet: {e}"}, status=500)


@require_GET
def zoom_in(request):
    global workbook
    if not workbook:
        return JsonResponse({"message": "No workbook loaded"}, status=400)
    try:
        workbook.Application.ActiveWindow.Zoom += 10
        return JsonResponse({"message": "Zoomed in."})
    except Exception as e:
        return JsonResponse({"message": f"Error zooming in: {e}"}, status=500)


@require_GET
def zoom_out(request):
    global workbook
    if not workbook:
        return JsonResponse({"message": "No workbook loaded"}, status=400)
    try:
        workbook.Application.ActiveWindow.Zoom -= 10
        return JsonResponse({"message": "Zoomed out."})
    except Exception as e:
        return JsonResponse({"message": f"Error zooming out: {e}"}, status=500)


@require_GET
def scroll_up(request):
    global workbook
    if not workbook:
        return JsonResponse({"message": "No workbook loaded"}, status=400)
    try:
        workbook.Application.ActiveWindow.ScrollRow -= 1
        return JsonResponse({"message": "Scrolled up."})
    except Exception as e:
        return JsonResponse({"message": f"Error scrolling up: {e}"}, status=500)


@require_GET
def scroll_down(request):
    global workbook
    if not workbook:
        return JsonResponse({"message": "No workbook loaded"}, status=400)
    try:
        workbook.Application.ActiveWindow.ScrollRow += 1
        return JsonResponse({"message": "Scrolled down."})
    except Exception as e:
        return JsonResponse({"message": f"Error scrolling down: {e}"}, status=500)


@require_GET
def scroll_left(request):
    global workbook
    if not workbook:
        return JsonResponse({"message": "No workbook loaded"}, status=400)
    try:
        workbook.Application.ActiveWindow.ScrollColumn -= 1
        return JsonResponse({"message": "Scrolled left."})
    except Exception as e:
        return JsonResponse({"message": f"Error scrolling left: {e}"}, status=500)


@require_GET
def scroll_right(request):
    global workbook
    if not workbook:
        return JsonResponse({"message": "No workbook loaded"}, status=400)
    try:
        workbook.Application.ActiveWindow.ScrollColumn += 1
        return JsonResponse({"message": "Scrolled right."})
    except Exception as e:
        return JsonResponse({"message": f"Error scrolling right: {e}"}, status=500)

@require_GET
def excel_controls(request):
    return JsonResponse({"message": "Excel controls loaded", "status": "success"})