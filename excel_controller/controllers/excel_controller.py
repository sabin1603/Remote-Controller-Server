from common.worker_thread import WorkerThread
import xlwings as xw
import os
from django.conf import settings


class ExcelWorkerThread(WorkerThread):
    """Dedicated worker thread for handling Excel operations."""

    def __init__(self):
        super().__init__(self.initialize_excel)

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

    def close_workbook(self):
        if self.worker_thread:
            self.worker_thread.add_to_queue(self.worker_thread.close_workbook)
            return True, "Workbook closed successfully."
        return False, "No active workbook to close."

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
        """Cleanup resources and stop the worker thread."""
        if self.worker_thread:
            self.worker_thread.add_to_queue(self.worker_thread.quit_excel)
            self.worker_thread.stop()
            self.worker_thread.join()
            self.worker_thread = None
