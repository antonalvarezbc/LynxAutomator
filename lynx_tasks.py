"""Background task state, with no dependency on Tk or any other UI toolkit."""
from queue import Empty, Queue
from threading import Event, Lock, Thread


class TaskCancelled(Exception):
    pass


class TaskContext:
    def __init__(self):
        self.stop = Event()
        self._lock = Lock()
        self._progress = (None, '')

    def checkpoint(self):
        if self.stop.is_set():
            raise TaskCancelled()

    def report(self, text, fraction=None):
        # Keep only the latest update; a long video must not fill the UI queue.
        with self._lock:
            self._progress = (fraction, str(text))

    def progress(self):
        with self._lock:
            return self._progress


class BackgroundTask:
    def __init__(self):
        self.context = None
        self.thread = None
        self._results = Queue()
        self.busy = False

    def start(self, work):
        if self.busy:
            raise RuntimeError('Another task is still running.')
        self.context = TaskContext()
        self.busy = True

        def run():
            try:
                self.context.checkpoint()
                result = work(self.context)
                outcome = ('success', result)
            except TaskCancelled:
                outcome = ('cancelled', None)
            except Exception as exc:
                outcome = ('error', exc)
            self._results.put(outcome)

        # Do not abandon an in-progress file write on interpreter shutdown.
        self.thread = Thread(target=run, name='LynxAutomator-worker', daemon=False)
        try:
            self.thread.start()
        except Exception:
            self.busy = False
            raise

    def cancel(self):
        if self.busy:
            self.context.stop.set()

    def take_result(self):
        try:
            result = self._results.get_nowait()
        except Empty:
            return None
        self.busy = False
        return result
