"""Worker chạy lời gọi mạng ở luồng nền để không treo giao diện."""
from PyQt6.QtCore import QThread, pyqtSignal

from ..api_client import ApiError

# Giữ tham chiếu các worker đang chạy để không bị thu gom rác giữa chừng
# (đặc biệt với worker "bắn rồi quên" khi cửa sổ gọi đã đóng).
_ALIVE: set["CallWorker"] = set()


class CallWorker(QThread):
    """Chạy `fn(*args, **kwargs)` ở nền, phát `ok`/`err`."""

    ok = pyqtSignal(object)
    err = pyqtSignal(str)

    def __init__(self, fn, *args, **kwargs):
        super().__init__()
        self._fn = fn
        self._args = args
        self._kwargs = kwargs
        _ALIVE.add(self)
        self.finished.connect(lambda: _ALIVE.discard(self))

    def run(self):
        try:
            result = self._fn(*self._args, **self._kwargs)
            self.ok.emit(result)
        except ApiError as e:
            self.err.emit(e.message)
        except Exception as e:  # noqa: BLE001 — mọi lỗi ngoài dự kiến
            self.err.emit(f"Lỗi không xác định: {e}")
