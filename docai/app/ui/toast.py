"""Toast đơn giản — báo số liệu prompt caching sau mỗi lần gọi Claude API."""
from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QLabel, QWidget

from ...logging_config import get_logger, log_call

logger = get_logger(__name__)

_DURATION_MS = 4000
_MARGIN = 16


class CacheToast(QLabel):
    """Banner nhỏ nổi ở góc dưới-phải cửa sổ, tự ẩn sau vài giây.

    `class` (cacheToastHit/New/None) đổi màu tuỳ cache có được đọc hay không —
    dùng selector .class trong QSS vì attribute selector không tô nền được
    trên bản Qt đang dùng (xem styles.py).
    """

    @log_call
    def __init__(self, parent: QWidget):
        super().__init__(parent)
        self.setObjectName("cacheToast")
        self.hide()
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self.hide)

    @log_call
    def show_usage(self, usage: dict | None):
        if not usage:
            return
        read = int(usage.get("cache_read_input_tokens") or 0)
        created = int(usage.get("cache_creation_input_tokens") or 0)
        fresh = int(usage.get("input_tokens") or 0)

        if read:
            css_class, headline = "cacheToastHit", "Cache HIT"
        elif created:
            css_class, headline = "cacheToastNew", "Cache mới ghi"
        else:
            css_class, headline = "cacheToastNone", "Không dùng cache"

        self.setProperty("class", css_class)
        self.style().unpolish(self)
        self.style().polish(self)

        self.setText(
            f"{headline} — đọc: {read:,} · ghi: {created:,} · mới: {fresh:,} token"
        )
        self.adjustSize()
        self.reposition()
        self.show()
        self.raise_()
        self._timer.start(_DURATION_MS)

    @log_call
    def reposition(self):
        parent = self.parentWidget()
        if not parent:
            return
        pos_x = parent.width() - self.width() - _MARGIN
        pos_y = parent.height() - self.height() - _MARGIN
        self.move(max(_MARGIN, pos_x), max(_MARGIN, pos_y))
