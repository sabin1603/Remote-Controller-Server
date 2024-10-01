from django.http import JsonResponse
from pywinauto import Application

from common.worker_thread import WorkerThread
import os
import urllib.parse
import traceback
import win32com.client
from django.conf import settings

class WordWorkerThread(WorkerThread):
    """Dedicated worker thread for handling Word operations."""

    def __init__(self):
        super().__init__(self.initialize_word)
        self.app = None
        self.document = None

    def initialize_word(self):
        """Initializes the Word application."""
        if self.app:
            try:
                self.quit_word()
            except Exception as e:
                print("Error: " + e)
        try:
            self.app = win32com.client.Dispatch("Word.Application")
            self.app.Visible = True
        except Exception as e:
            print(f"Failed to initialize Word: {e}")
            self.app = None

    def bring_to_front(self):
        try:
            # Connect to the running Word application
            app = Application().connect(path="WINWORD.EXE")
            app.top_window().set_focus()  # Set focus to the top window (Word)
            return JsonResponse({'status': 'success', 'message': 'Word brought to front successfully.'})
        except Exception as e:
            return JsonResponse({'status': 'error', 'message': f'Failed to bring Word to front: {e}'}, status=500)

    def open_document(self, file_path):
        """Opens a Word document."""

        file_path = urllib.parse.unquote(file_path)
        print(f"Original file path: {file_path}")

        # Join and normalize the path
        full_path = os.path.normpath(os.path.join(settings.ROOT_DIR, file_path))
        print(f"Normalized file path: {full_path}")

        if not os.path.isfile(full_path):
            print(f"File does not exist: {full_path}")
            return False, "File does not exist."

        try:
            print(f"Attempting to open document at: {full_path}")
            if self.document:
                self.document.Close()
            self.document = self.app.Documents.Open(full_path)
        except Exception as e:
            error_msg = f"Failed to open document: {e}"
            traceback_msg = traceback.format_exc()
            print(error_msg)
            print(traceback_msg)
            self.document = None
            raise Exception(error_msg)

    def close_document(self):
        """Closes the Word document and logs any errors."""
        if self.document:
            try:
                self.document.Close()
                self.document = None
            except Exception as e:
                print(f"Failed to close document: {e}")
                # Retry or cleanup resources if needed
                self.cleanup()

    def scroll_up(self):
        """Scrolls up in the Word document."""
        if self.document:
            try:
                if self.app.ActiveWindow.View.ReadingLayout:
                    self.app.ActiveWindow.LargeScroll(Down=-1)
                else:
                    self.app.ActiveWindow.SmallScroll(Down=-1)
            except Exception as e:
                print(f"Failed to scroll up: {e}")

    def scroll_down(self):
        """Scrolls down in the Word document."""
        if self.document:
            try:
                if self.app.ActiveWindow.View.ReadingLayout:
                    self.app.ActiveWindow.LargeScroll(Down=1)
                else:
                    self.app.ActiveWindow.SmallScroll(Down=1)
            except Exception as e:
                print(f"Failed to scroll down: {e}")

    def zoom_in(self):
        """Zooms in the Word document."""
        if self.document:
            try:
                self.app.ActiveWindow.View.Zoom.Percentage += 10
            except Exception as e:
                print(f"Failed to zoom in: {e}")

    def zoom_out(self):
        """Zooms out the Word document."""
        if self.document:
            try:
                self.app.ActiveWindow.View.Zoom.Percentage -= 10
            except Exception as e:
                print(f"Failed to zoom out: {e}")

    def enable_read_mode(self):
        """Enables read mode in the Word document."""
        if self.document:
            try:
                self.app.ActiveWindow.View.ReadingLayout = True
            except Exception as e:
                print(f"Failed to enable read mode: {e}")

    def disable_read_mode(self):
        """Disables read mode in the Word document."""
        if self.document:
            try:
                self.app.ActiveWindow.View.ReadingLayout = False
            except Exception as e:
                print(f"Failed to disable read mode: {e}")

    def quit_word(self):
        """Quits the Word application."""
        if self.app:
            try:
                self.app.Quit()
                self.app = None
            except Exception as e:
                print(f"Failed to quit Word: {e}")


class WordController:
    def __init__(self):
        self.worker_thread = None

    def start_worker(self):
        """Starts the worker thread if not already running."""
        if not self.worker_thread:
            self.worker_thread = WordWorkerThread()
            self.worker_thread.start()

    def open_document(self, file_path):
        self.start_worker()
        self.worker_thread.add_to_queue(self.worker_thread.open_document, file_path)
        return True, "Document loaded successfully."

    def bring_to_front(self):
        """Brings the Word application to the front."""
        if self.worker_thread:
            self.worker_thread.add_to_queue(self.worker_thread.bring_to_front)
            return True, "Word brought to front."
        return False, "Word application is not running."

    def close_document(self):
        if self.worker_thread:
            self.worker_thread.add_to_queue(self.worker_thread.close_document)
            self.cleanup()
            return True, "Document closed successfully."
        return False, "No active document to close."

    def scroll_up(self):
        self.worker_thread.add_to_queue(self.worker_thread.scroll_up)
        return True, "Scrolled up."

    def scroll_down(self):
        self.worker_thread.add_to_queue(self.worker_thread.scroll_down)
        return True, "Scrolled down."

    def zoom_in(self):
        self.worker_thread.add_to_queue(self.worker_thread.zoom_in)
        return True, "Zoomed in."

    def zoom_out(self):
        self.worker_thread.add_to_queue(self.worker_thread.zoom_out)
        return True, "Zoomed out."

    def enable_read_mode(self):
        self.worker_thread.add_to_queue(self.worker_thread.enable_read_mode)
        return True, "Read mode enabled."

    def disable_read_mode(self):
        self.worker_thread.add_to_queue(self.worker_thread.disable_read_mode)
        return True, "Read mode disabled."

    def cleanup(self):
        """Cleanup resources and stop the worker thread."""
        if self.worker_thread:
            self.worker_thread.add_to_queue(self.worker_thread.quit_word)  
            self.worker_thread.add_to_queue(self.worker_thread.stop)  
            self.worker_thread.join()  
            self.worker_thread = None  
