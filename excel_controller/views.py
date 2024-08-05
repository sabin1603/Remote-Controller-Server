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

# Global Excel object and workbook
excel_app = None
workbook = None
current_workbook = None

def get_excel_instance():
    global excel_app
    if not excel_app:
        try:
            initialize_com()
            excel_app = win32com.client.Dispatch("Excel.Application")
            excel_app.Visible = True
        except Exception as e:
            return None, str(e)
    return excel_app, None

@require_GET
def list_workbooks(request):
    xlsx_files = []
    for root, _, files in os.walk(settings.ROOT_EXCEL_DIR):
        for file in files:
            if file.endswith('.xlsx') or file.endswith('.xls'):
                xlsx_files.append(os.path.join(root, file))
    return JsonResponse({'workbooks': xlsx_files})

@require_GET
def open_workbook(request, file_name):
    global workbook, current_workbook
    app, error = get_excel_instance()
    if error:
        return JsonResponse({"message": f"Failed to access Excel: {error}"}, status=500)

    full_path = os.path.join(settings.ROOT_EXCEL_DIR, file_name)

    try:
        if workbook:
            workbook.Close(False)
        workbook = app.Workbooks.Open(full_path)
        current_workbook = full_path
        return JsonResponse({"message": "Workbook loaded successfully.", "file_name": file_name})
    except Exception as e:
        return JsonResponse({"message": f"Failed to load workbook: {e}"}, status=500)

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
        workbook.ActiveWindow.Zoom += 10
        return JsonResponse({"message": "Zoomed in."})
    except Exception as e:
        return JsonResponse({"message": f"Error zooming in: {e}"}, status=500)

@require_GET
def zoom_out(request):
    global workbook
    if not workbook:
        return JsonResponse({"message": "No workbook loaded"}, status=400)
    try:
        workbook.ActiveWindow.Zoom -= 10
        return JsonResponse({"message": "Zoomed out."})
    except Exception as e:
        return JsonResponse({"message": f"Error zooming out: {e}"}, status=500)

@require_GET
def scroll_up(request):
    global workbook
    if not workbook:
        return JsonResponse({"message": "No workbook loaded"}, status=400)
    try:
        workbook.ActiveWindow.ScrollRow -= 1
        return JsonResponse({"message": "Scrolled up."})
    except Exception as e:
        return JsonResponse({"message": f"Error scrolling up: {e}"}, status=500)

@require_GET
def scroll_down(request):
    global workbook
    if not workbook:
        return JsonResponse({"message": "No workbook loaded"}, status=400)
    try:
        workbook.ActiveWindow.ScrollRow += 1
        return JsonResponse({"message": "Scrolled down."})
    except Exception as e:
        return JsonResponse({"message": f"Error scrolling down: {e}"}, status=500)

@require_GET
def scroll_left(request):
    global workbook
    if not workbook:
        return JsonResponse({"message": "No workbook loaded"}, status=400)
    try:
        workbook.ActiveWindow.ScrollColumn -= 1
        return JsonResponse({"message": "Scrolled left."})
    except Exception as e:
        return JsonResponse({"message": f"Error scrolling left: {e}"}, status=500)

@require_GET
def scroll_right(request):
    global workbook
    if not workbook:
        return JsonResponse({"message": "No workbook loaded"}, status=400)
    try:
        workbook.ActiveWindow.ScrollColumn += 1
        return JsonResponse({"message": "Scrolled right."})
    except Exception as e:
        return JsonResponse({"message": f"Error scrolling right: {e}"}, status=500)
