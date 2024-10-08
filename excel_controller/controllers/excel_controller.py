from django.http import JsonResponse
from common.worker_thread import WorkerThread
import xlwings as xw
import os
from django.conf import settings
from pywinauto import application

class ExcelWorkerThread(WorkerThread):
    """Dedicated worker thread for handling Excel operations."""

    def __init__(self):
        super().__init__(self.initialize_excel)

    def initialize_excel(self):
        """Initializes the Excel application."""
        if self.app:
            try:
                self.quit_excel()
            except Exception as e:
                print("Error: " + e)
        try:
            self.app = xw.App(visible=True)  # Open Excel in visible mode
            self.app.display_alerts = False  # Prevent Excel dialogs
        except Exception as e:
            print(f"Failed to initialize Excel: {e}")
            self.app = None

    def open_workbook(self, file_path):
        """Opens an Excel workbook."""
        full_path = os.path.join(settings.ROOT_EXCEL_DIR, file_path)
        self.workbook = self.app.books.open(full_path)

    def bring_to_front(self):
        try:
            app = application.Application().connect(path="EXCEL.EXE")
            app.top_window().set_focus()
            return JsonResponse({'status': 'success', 'message': 'Excel brought to front successfully.'})
        except Exception as e:
            return JsonResponse({'status': 'error', 'message': f'Failed to bring Excel to front: {e}'}, status=500)

    def close_workbook(self):
        """Closes the Excel workbook and logs any errors."""
        if self.workbook:
            try:
                self.workbook.close()
                self.workbook = None
            except Exception as e:
                print(f"Failed to close workbook: {e}")
                # Retry or cleanup resources if needed
                self.cleanup()

    def quit_excel(self):
        """Quits the Excel application."""
        if self.app:
            try:
                self.app.quit()
                self.app = None
            except Exception as e:
                print(f"Failed to quit Excel: {e}")


class ExcelController:
    def __init__(self):
        self.worker_thread = None

    def start_worker(self):
        """Starts the worker thread if not already running."""
        if not self.worker_thread:
            self.worker_thread = ExcelWorkerThread()
            self.worker_thread.start()

    def open_workbook(self, file_path):
        self.start_worker()
        self.worker_thread.add_to_queue(self.worker_thread.open_workbook, file_path)
        return True, "Workbook loaded successfully."

    def bring_to_front(self):
        """Brings the Excel application to the front."""
        if self.worker_thread:
            self.worker_thread.add_to_queue(self.worker_thread.bring_to_front)
            return True, "Excel brought to front."
        return False, "Excel application is not running."

    def close_workbook(self):
        """Closes the currently open workbook and cleans up resources."""
        if self.worker_thread:
            self.worker_thread.add_to_queue(self.worker_thread.close_workbook)
            self.cleanup()  # Call cleanup after closing the workbook
            return True, "Workbook closed successfully."
        return False, "No active workbook to close."

    def change_worksheet(self, next_sheet=True):
        def change_sheet():
            if not self.worker_thread.workbook:
                raise Exception("No workbook loaded.")

            workbook = self.worker_thread.workbook
            current_sheet_index = workbook.api.ActiveSheet.Index

            if next_sheet:
                new_index = current_sheet_index + 1
                if new_index <= workbook.api.Sheets.Count:
                    workbook.api.Sheets(new_index).Activate()
                else:
                    raise IndexError("Already on the last worksheet.")
            else:
                new_index = current_sheet_index - 1
                if new_index >= 1:
                    workbook.api.Sheets(new_index).Activate()
                else:
                    raise IndexError("Already on the first worksheet.")

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
        """Cleanup resources and stop the worker thread."""
        if self.worker_thread:
            self.worker_thread.add_to_queue(self.worker_thread.quit_excel)  
            self.worker_thread.add_to_queue(self.worker_thread.stop)  
            self.worker_thread.join()  
            self.worker_thread = None  

