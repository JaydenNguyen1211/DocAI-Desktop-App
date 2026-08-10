"""Worker chạy lời gọi mạng ở luồng nền để không treo giao diện."""
from PySide6.QtCore import QThread, Signal

from ..account.api_client import ApiError

from ..logging_config import get_logger, log_call

logger = get_logger(__name__)

# Giữ tham chiếu các worker đang chạy để không bị thu gom rác giữa chừng
# (đặc biệt với worker "bắn rồi quên" khi cửa sổ gọi đã đóng).
_ALIVE: set["CallWorker"] = set()


class CallWorker(QThread):
    """Chạy `fn(*args, **kwargs)` ở nền, phát `ok`/`err`."""

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
        except Exception as exc:  # noqa: BLE001 — mọi lỗi ngoài dự kiến
            logger.exception("Background call %r failed unexpectedly", self._fn)
            self.err.emit(f"Lỗi không xác định: {exc}")
