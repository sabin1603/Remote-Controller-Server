import threading
import pythoncom

class WorkerThread(threading.Thread):
    """Generic worker thread for handling application operations."""

    def __init__(self, init_func=None):
        super().__init__()
        self.lock = threading.Lock()
        self.queue = []
        self.running = True
        self.init_func = init_func
        self.app = None
        self.workbook = None  # or document, presentation, etc. for other modules

    def run(self):
        """Thread main loop to process commands."""
        pythoncom.CoInitialize()
        if self.init_func:
            self.init_func()
        try:
            while self.running:
                if self.queue:
                    with self.lock:
                        func, args, kwargs = self.queue.pop(0)
                    try:
                        func(*args, **kwargs)
                    except Exception as e:
                        print(f"Error during operation: {e}")
        finally:
            pythoncom.CoUninitialize()

    def add_to_queue(self, func, *args, **kwargs):
        """Add a function to the queue for execution in the thread."""
        with self.lock:
            self.queue.append((func, args, kwargs))

    def stop(self):
        """Stops the worker thread."""
        self.running = False
        self.add_to_queue(lambda: None)  # Add a no-op to unblock the loop
