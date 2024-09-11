from common.worker_thread import WorkerThread
import os
import urllib.parse
import subprocess
import traceback
import pyautogui

ADOBE_ACROBAT_PATH = r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"

class PDFWorkerThread(WorkerThread):
    """Dedicated worker thread for handling PDF operations."""

    def __init__(self):
        super().__init__(self.initialize_acrobat)
        self.process = None

    def initialize_acrobat(self):
        """Initializes the Adobe Acrobat application."""
        if self.app:
            try:
                self.quit_acrobat
            except Exception as e:
                print("Error: " + e)
        try:
            # Start a new Adobe Acrobat instance
            self.process = subprocess.Popen([ADOBE_ACROBAT_PATH], shell=True)
            print("Adobe Acrobat initialized successfully.")
        except Exception as e:
            print(f"Failed to initialize Adobe Acrobat: {e}")
            self.process = None

    def open_pdf(self, file_path):
        """Opens a PDF."""
        file_path = urllib.parse.unquote(file_path)
        print(f"Original file path: {file_path}")

        # Normalize the provided file path
        full_path = os.path.normpath(file_path)  # Since file_path is already absolute, use it directly
        print(f"Normalized file path: {full_path}")

        if not os.path.isfile(full_path):
            print(f"File does not exist: {full_path}")
            return False, "File does not exist."

        try:
            print(f"Attempting to open PDF at: {full_path}")
            if self.process:
                self.close_pdf()

            # Correct the reference to the Acrobat path
            self.process = subprocess.Popen([ADOBE_ACROBAT_PATH, full_path],
                                            shell=False)  # Use shell=False for better handling
            return True, "PDF loaded successfully."
        except Exception as e:
            error_msg = f"Failed to open PDF: {e}"
            traceback_msg = traceback.format_exc()
            print(error_msg)
            print(traceback_msg)
            self.process = None
            raise Exception(error_msg)

    def close_pdf(self):
        """Closes the PDF document."""
        if self.process:
            try:
                self.quit_acrobat()
                self.process = None
            except Exception as e:
                print(f"Failed to close PDF: {e}")
                self.cleanup()

    def quit_acrobat(self):
        """Quits the Adobe Acrobat application."""
        if self.process:
            try:
                self.process.terminate()
                self.process.wait()
                self.process = None
                print("Adobe Acrobat process terminated.")
            except Exception as e:
                print(f"Failed to quit Adobe Acrobat: {e}")

    def scroll_up(self):
        """Scrolls up in the PDF document."""
        pyautogui.scroll(200)

    def scroll_down(self):
        """Scrolls down in the PDF document."""
        pyautogui.scroll(-200)

    def zoom_in(self):
        """Zooms in the PDF document."""
        pyautogui.hotkey('ctrl', '=')

    def zoom_out(self):
        """Zooms out the PDF document."""
        pyautogui.hotkey('ctrl', '-')

    def enable_read_mode(self):
        """Enables read mode in the PDF document."""
        pyautogui.hotkey('ctrl', 'h')  # Adjust the hotkey based on Acrobat's shortcuts

    def disable_read_mode(self):
        """Disables read mode in the PDF document."""
        pyautogui.hotkey('ctrl', 'h')  # Toggle to disable read mode

    def save_pdf(self):
        """Saves the PDF document."""
        pyautogui.hotkey('ctrl', 's')

    def print_pdf(self):
        """Opens the print dialog for the PDF document."""
        pyautogui.hotkey('ctrl', 'p')


class PDFController:
    def __init__(self):
        self.worker_thread = None

    def start_worker(self):
        """Starts the worker thread if not already running."""
        if not self.worker_thread:
            self.worker_thread = PDFWorkerThread()
            self.worker_thread.start()

    def open_pdf(self, file_path):
        self.start_worker()
        self.worker_thread.add_to_queue(self.worker_thread.open_pdf, file_path)
        return True, "PDF loaded successfully."

    def close_pdf(self):
        if self.worker_thread:
            self.worker_thread.add_to_queue(self.worker_thread.close_pdf)
            self.cleanup()
            return True, "PDF closed successfully."
        return False, "No active PDF to close."

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

    def save_pdf(self):
        self.worker_thread.add_to_queue(self.worker_thread.save_pdf)
        return True, "PDF saved."

    def print_pdf(self):
        self.worker_thread.add_to_queue(self.worker_thread.print_pdf)
        return True, "Print dialog opened."

    def cleanup(self):
        """Cleanup resources and stop the worker thread."""
        if self.worker_thread:
            self.worker_thread.add_to_queue(self.worker_thread.quit_acrobat)
            self.worker_thread.stop()
            self.worker_thread.join()
            self.worker_thread = None
