"""Worker that runs network calls on a background thread so the UI doesn't freeze."""
from PySide6.QtCore import QThread, Signal

from ..account.api_client import ApiError

from ..logging_config import get_logger, log_call

logger = get_logger(__name__)

# Keep a reference to running workers so they aren't garbage-collected
# mid-flight (especially "fire and forget" workers whose calling window has
# already closed).
_ALIVE: set["CallWorker"] = set()


class CallWorker(QThread):
    """Runs `fn(*args, **kwargs)` in the background, emits `ok`/`err`."""

    ok = Signal(object)
    err = Signal(str)

    @log_call
    def __init__(self, fn, *args, **kwargs):
        super().__init__()
        self._fn = fn
        self._args = args
        self._kwargs = kwargs
        _ALIVE.add(self)
        self.finished.connect(lambda: _ALIVE.discard(self))

    @log_call
    def run(self):
        try:
            result = self._fn(*self._args, **self._kwargs)
            self.ok.emit(result)
        except ApiError as exc:
            logger.warning(
                "Background call %r failed with ApiError: %s (code=%s)",
                self._fn, exc.message, exc.code)
            self.err.emit(exc.message)
        except Exception as exc:  # noqa: BLE001 — any unforeseen error
            logger.exception("Background call %r failed unexpectedly", self._fn)
            self.err.emit(f"Lỗi không xác định: {exc}")
