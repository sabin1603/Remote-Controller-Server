import traceback
import urllib.parse
from common.worker_thread import WorkerThread
import os
import win32com.client
from django.conf import settings

class PowerPointWorkerThread(WorkerThread):
    """Dedicated worker thread for handling PowerPoint operations."""

    def __init__(self):
        super().__init__(self.initialize_powerpoint)
        self.app = None
        self.presentation = None

    def initialize_powerpoint(self):
        """Initializes the PowerPoint application."""
        if self.app:
            try:
                self.quit_powerpoint()
            except Exception as e:
                print("Error: "+ e)
        try:
            self.app = win32com.client.Dispatch("PowerPoint.Application")
            self.app.Visible = True
        except Exception as e:
            print(f"Failed to initialize PowerPoint: {e}")
            self.app = None

    def open_presentation(self, file_path):
        """Opens a PowerPoint presentation."""
        file_path = urllib.parse.unquote(file_path)
        print(f"Original file path: {file_path}")

        # Join and normalize the path
        full_path = os.path.normpath(os.path.join(settings.ROOT_DIR, file_path))
        print(f"Normalized file path: {full_path}")

        if not os.path.isfile(full_path):
            print(f"File does not exist: {full_path}")
            return False, "File does not exist."

        try:
            print(f"Attempting to open presentation at: {full_path}")
            if self.presentation:
                self.presentation.Close()
            self.presentation = self.app.Presentations.Open(full_path)
        except Exception as e:
            error_msg = f"Failed to open presentation: {e}"
            traceback_msg = traceback.format_exc()
            print(error_msg)
            print(traceback_msg)
            self.presentation = None
            raise Exception(error_msg)

    def close_presentation(self):
        """Closes the PowerPoint presentation."""
        if self.presentation:
            try:
                self.presentation.Close()
                self.presentation = None
            except Exception as e:
                print(f"Failed to close presentation: {e}")
                self.cleanup()

    def quit_powerpoint(self):
        """Quits the PowerPoint application."""
        if self.app:
            try:
                self.app.Quit()
                self.app = None
            except Exception as e:
                print(f"Failed to quit PowerPoint: {e}")

    def next_slide(self):
        """Moves to the next slide in the presentation."""
        if self.presentation and self.presentation.SlideShowWindow:
            try:
                self.presentation.SlideShowWindow.View.Next()
            except Exception as e:
                print(f"Failed to move to next slide: {e}")

    def prev_slide(self):
        """Moves to the previous slide in the presentation."""
        if self.presentation and self.presentation.SlideShowWindow:
            try:
                self.presentation.SlideShowWindow.View.Previous()
            except Exception as e:
                print(f"Failed to move to previous slide: {e}")

    def start_presentation(self):
        """Starts the slideshow mode."""
        if self.presentation:
            try:
                self.presentation.SlideShowSettings.Run()
            except Exception as e:
                print(f"Failed to start presentation: {e}")

    def end_presentation(self):
        """Ends the slideshow mode."""
        if self.presentation and self.presentation.SlideShowWindow:
            try:
                self.presentation.SlideShowWindow.View.Exit()
            except Exception as e:
                print(f"Failed to end presentation: {e}")

class PowerPointController:
    def __init__(self):
        self.worker_thread = None

    def start_worker(self):
        """Starts the worker thread if not already running."""
        if not self.worker_thread:
            self.worker_thread = PowerPointWorkerThread()
            self.worker_thread.start()

    def open_presentation(self, file_path):
        self.start_worker()
        self.worker_thread.add_to_queue(self.worker_thread.open_presentation, file_path)
        return True, "Presentation loaded successfully."

    def close_presentation(self):
        if self.worker_thread:
            self.worker_thread.add_to_queue(self.worker_thread.close_presentation)
            self.cleanup()
            return True, "Presentation closed successfully."
        return False, "No active presentation to close."

    def next_slide(self):
        self.worker_thread.add_to_queue(self.worker_thread.next_slide)
        return True, "Moved to next slide."

    def prev_slide(self):
        self.worker_thread.add_to_queue(self.worker_thread.prev_slide)
        return True, "Moved to previous slide."

    def start_presentation(self):
        self.worker_thread.add_to_queue(self.worker_thread.start_presentation)
        return True, "Presentation started."

    def end_presentation(self):
        self.worker_thread.add_to_queue(self.worker_thread.end_presentation)
        return True, "Presentation ended."

    def cleanup(self):
        """Cleanup resources and stop the worker thread."""
        if self.worker_thread:
            self.worker_thread.add_to_queue(self.worker_thread.quit_powerpoint)
            self.worker_thread.add_to_queue(self.worker_thread.stop)  
            self.worker_thread.join()  
            self.worker_thread = None  
