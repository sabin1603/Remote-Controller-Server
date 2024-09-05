# views.py
import os
import xlwings as xw
from django.http import JsonResponse
from django.views.decorators.http import require_GET
from django.conf import settings

class ExcelController:
    def __init__(self):
        self.app = None
        self.workbook = None

    def open_workbook(self, file_path):
        try:
            full_path = os.path.join(settings.ROOT_EXCEL_DIR, file_path)
            self.app = xw.App(visible=True)  # Open Excel in visible mode
            self.workbook = self.app.books.open(full_path)
            return True, "Workbook loaded successfully."
        except Exception as e:
            return False, f"Failed to load workbook: {e}"

    def close_workbook(self):
        try:
            if self.workbook:
                self.workbook.close()
                self.workbook = None
            if self.app:
                self.app.quit()
                self.app = None
            return True, "Workbook closed successfully."
        except Exception as e:
            return False, f"Error closing workbook: {e}"

    def change_worksheet(self, next_sheet=True):
        try:
            if not self.workbook:
                return False, "No workbook loaded."
            current_sheet_index = self.workbook.sheets.active.index
            new_index = current_sheet_index + 1 if next_sheet else current_sheet_index - 1
            if 1 <= new_index <= len(self.workbook.sheets):
                self.workbook.sheets[new_index].activate()
                return True, "Worksheet changed."
            else:
                return False, "Already on the last or first worksheet."
        except Exception as e:
            return False, f"Error changing worksheet: {e}"

    def zoom(self, zoom_in=True):
        try:
            if not self.workbook:
                return False, "No workbook loaded."
            active_window = self.workbook.app.api.ActiveWindow
            zoom_value = active_window.Zoom
            new_zoom = zoom_value + 10 if zoom_in else zoom_value - 10
            active_window.Zoom = new_zoom
            return True, "Zoom adjusted."
        except Exception as e:
            return False, f"Error adjusting zoom: {e}"

    def scroll(self, direction):
        try:
            if not self.workbook:
                return False, "No workbook loaded."
            active_window = self.workbook.app.api.ActiveWindow
            if direction == 'up':
                active_window.ScrollRow -= 1
            elif direction == 'down':
                active_window.ScrollRow += 1
            elif direction == 'left':
                active_window.ScrollColumn -= 1
            elif direction == 'right':
                active_window.ScrollColumn += 1
            return True, "Scrolled successfully."
        except Exception as e:
            return False, f"Error scrolling: {e}"

excel_controller = ExcelController()

@require_GET
def open_workbook(request, file_path):
    success, message = excel_controller.open_workbook(file_path)
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def close_workbook(request):
    success, message = excel_controller.close_workbook()
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def next_worksheet(request):
    success, message = excel_controller.change_worksheet(next_sheet=True)
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def previous_worksheet(request):
    success, message = excel_controller.change_worksheet(next_sheet=False)
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def zoom_in(request):
    success, message = excel_controller.zoom(zoom_in=True)
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def zoom_out(request):
    success, message = excel_controller.zoom(zoom_in=False)
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def scroll_up(request):
    success, message = excel_controller.scroll('up')
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def scroll_down(request):
    success, message = excel_controller.scroll('down')
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def scroll_left(request):
    success, message = excel_controller.scroll('left')
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def scroll_right(request):
    success, message = excel_controller.scroll('right')
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def excel_controls(request):
    return JsonResponse({"message": "Excel controls loaded", "status": "success"})
