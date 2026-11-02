import logging

from PySide6.QtCore import QThread, Signal


class TaskWorker(QThread):
    succeeded = Signal(object)
    failed = Signal(str)

    def __init__(self, operation, parent=None):
        super().__init__(parent)
        self.operation = operation

    def run(self):
        try:
            result = self.operation()
        except Exception as exc:
            logging.getLogger(__name__).exception("Asztali háttérművelet sikertelen")
            self.failed.emit(str(exc))
        else:
            self.succeeded.emit(result)
