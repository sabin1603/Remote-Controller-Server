import pythoncom
import os
import xlwings as xw
from django.http import JsonResponse
from django.views.decorators.http import require_GET
from django.conf import settings
import threading

class ExcelWorkerThread(threading.Thread):
    """Dedicated worker thread for handling Excel operations."""
    def __init__(self):
        super().__init__()
        self.app = None
        self.workbook = None
        self.lock = threading.Lock()
        self.queue = []
        self.running = True

    def run(self):
        """Thread main loop to process Excel commands."""
        pythoncom.CoInitialize()
        try:
            while self.running:
                if self.queue:
                    with self.lock:
                        func, args, kwargs = self.queue.pop(0)
                    try:
                        func(*args, **kwargs)
                    except Exception as e:
                        print(f"Error during Excel operation: {e}")
        finally:
            pythoncom.CoUninitialize()

    def add_to_queue(self, func, *args, **kwargs):
        """Add a function to the queue for execution in the Excel thread."""
        with self.lock:
            self.queue.append((func, args, kwargs))

    def initialize_excel(self):
        """Initializes the Excel application."""
        self.app = xw.App(visible=True)  # Open Excel in visible mode

    def open_workbook(self, file_path):
        """Opens an Excel workbook."""
        full_path = os.path.join(settings.ROOT_EXCEL_DIR, file_path)
        self.workbook = self.app.books.open(full_path)

    def close_workbook(self):
        """Closes the Excel workbook."""
        if self.workbook:
            self.workbook.close()
            self.workbook = None

    def quit_excel(self):
        """Quits the Excel application."""
        if self.app:
            self.app.quit()
            self.app = None

    def stop(self):
        """Stops the worker thread."""
        self.running = False
        self.add_to_queue(lambda: None)  # Add a no-op to unblock the loop

class ExcelController:
    def __init__(self):
        self.worker_thread = ExcelWorkerThread()
        self.worker_thread.start()
        self.worker_thread.add_to_queue(self.worker_thread.initialize_excel)

    def open_workbook(self, file_path):
        self.worker_thread.add_to_queue(self.worker_thread.open_workbook, file_path)
        return True, "Workbook loaded successfully."

    def close_workbook(self):
        self.worker_thread.add_to_queue(self.worker_thread.close_workbook)
        return True, "Workbook closed successfully."

    def change_worksheet(self, next_sheet=True):
        def change_sheet():
            if not self.worker_thread.workbook:
                raise Exception("No workbook loaded.")
            current_sheet_index = self.worker_thread.workbook.api.ActiveSheet.Index
            new_index = current_sheet_index + 1 if next_sheet else current_sheet_index - 1
            sheet_count = self.worker_thread.workbook.sheets.count
            if 1 <= new_index <= sheet_count:
                self.worker_thread.workbook.sheets[new_index].activate()
        self.worker_thread.add_to_queue(change_sheet)
        return True, "Worksheet changed."

    def zoom(self, zoom_in=True):
        def adjust_zoom():
            if not self.worker_thread.workbook:
                raise Exception("No workbook loaded.")
            active_window = self.worker_thread.workbook.app.api.ActiveWindow
            zoom_value = active_window.Zoom
            new_zoom = zoom_value + 10 if zoom_in else zoom_value - 10
            active_window.Zoom = new_zoom
        self.worker_thread.add_to_queue(adjust_zoom)
        return True, "Zoom adjusted."

    def scroll(self, direction):
        def perform_scroll():
            if not self.worker_thread.workbook:
                raise Exception("No workbook loaded.")
            active_window = self.worker_thread.workbook.app.api.ActiveWindow
            if direction == 'up':
                active_window.ScrollRow -= 1
            elif direction == 'down':
                active_window.ScrollRow += 1
            elif direction == 'left':
                active_window.ScrollColumn -= 1
            elif direction == 'right':
                active_window.ScrollColumn += 1
        self.worker_thread.add_to_queue(perform_scroll)
        return True, "Scrolled successfully."

    def cleanup(self):
        self.worker_thread.add_to_queue(self.worker_thread.quit_excel)
        self.worker_thread.stop()
        self.worker_thread.join()

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
