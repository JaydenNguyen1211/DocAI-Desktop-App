"""Vùng nội dung trung tâm cho chế độ Thư mục — 3 trạng thái:
trống (4.1) → đang quét & lập chỉ mục (4.3) → sẵn sàng (4.4)."""
from PySide6.QtCore import Qt, QSize, Signal
from PySide6.QtWidgets import (
    QFrame, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QStackedWidget,
    QProgressBar, QWidget,
)

from ..constants import FILE_BADGE_STYLE
from .icons import folder_icon

_EMPTY, _SCANNING, _READY = 0, 1, 2


class _CountBox(QFrame):
    def __init__(self, count: int, label: str, parent=None):
        super().__init__(parent)
        self.setObjectName("countBox")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(14, 9, 14, 9)
        lay.setSpacing(2)
        num = QLabel(str(count))
        num.setObjectName("countBoxNum")
        num.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(num)
        lbl = QLabel(label)
        lbl.setObjectName("countBoxLabel")
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(lbl)


class FolderWorkspace(QFrame):
    """Panel bên phải khi tab Thư mục đang chọn — thay cho preview/chat trống."""

    open_requested = Signal()          # bấm "Mở thư mục làm việc"
    select_all_requested = Signal()    # bấm "Chọn tất cả"
    start_chat_requested = Signal()    # bấm "Bắt đầu chat"

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("centralChat")

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        self._stack = QStackedWidget()
        root.addWidget(self._stack)

        self._stack.addWidget(self._build_empty())
        self._stack.addWidget(self._build_scanning())
        self._stack.addWidget(self._build_ready())

    # ── 4.1 — trống ─────────────────────────────────────────────────────────

    def _build_empty(self) -> QWidget:
        page = QWidget()
        lay = QVBoxLayout(page)
        lay.setContentsMargins(48, 40, 48, 40)
        lay.addStretch(3)

        col = QVBoxLayout()
        col.setSpacing(0)

        icon = QLabel()
        icon.setObjectName("folderEmptyIcon")
        icon.setFixedSize(52, 52)
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon.setPixmap(folder_icon("#2A6FDB", 24).pixmap(QSize(24, 24)))
        col.addWidget(icon, alignment=Qt.AlignmentFlag.AlignHCenter)
        col.addSpacing(18)

        title = QLabel("Làm việc với cả một thư mục")
        title.setObjectName("welcomeTitle")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        col.addWidget(title)
        col.addSpacing(8)

        sub = QLabel(
            "Mở thư mục hóa đơn / chứng từ, DocAI sẽ lập chỉ mục để bạn\n"
            "tick chọn file đưa vào ngữ cảnh chat.")
        sub.setObjectName("welcomeSub")
        sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        col.addWidget(sub)
        col.addSpacing(20)

        btn = QPushButton("  Mở thư mục làm việc")
        btn.setProperty("class", "primaryBtn")
        btn.setIcon(folder_icon("#FFFFFF"))
        btn.setIconSize(QSize(15, 15))
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.clicked.connect(self.open_requested.emit)
        col.addWidget(btn, alignment=Qt.AlignmentFlag.AlignHCenter)

        lay.addLayout(col)
        lay.addStretch(4)
        return page

    # ── 4.3 — đang quét & lập chỉ mục ───────────────────────────────────────

    def _build_scanning(self) -> QWidget:
        page = QWidget()
        lay = QVBoxLayout(page)
        lay.setContentsMargins(28, 28, 28, 28)
        lay.setSpacing(18)
        lay.addStretch(1)

        self._scan_title = QLabel("Đang lập chỉ mục…")
        self._scan_title.setObjectName("welcomeTitle")
        self._scan_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(self._scan_title)

        bar_row = QHBoxLayout()
        bar_row.addStretch()
        self._scan_bar = QProgressBar()
        self._scan_bar.setObjectName("folderScanBar")
        self._scan_bar.setFixedWidth(320)
        self._scan_bar.setFixedHeight(6)
        self._scan_bar.setTextVisible(False)
        self._scan_bar.setMaximum(100)
        bar_row.addWidget(self._scan_bar)
        bar_row.addStretch()
        lay.addLayout(bar_row)

        self._count_row = QHBoxLayout()
        self._count_row.setSpacing(10)
        count_wrap = QHBoxLayout()
        count_wrap.addStretch()
        count_wrap.addLayout(self._count_row)
        count_wrap.addStretch()
        lay.addLayout(count_wrap)

        hint = QLabel("Bỏ qua file không hỗ trợ · không tải nội dung cho tới khi bạn tick")
        hint.setObjectName("footnote")
        hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(hint)

        lay.addStretch(1)
        return page

    # ── 4.4 — sẵn sàng ──────────────────────────────────────────────────────

    def _build_ready(self) -> QWidget:
        page = QWidget()
        lay = QVBoxLayout(page)
        lay.setContentsMargins(48, 40, 48, 40)
        lay.addStretch(3)

        col = QVBoxLayout()
        col.setSpacing(0)

        check = QLabel("✓")
        check.setObjectName("folderReadyIcon")
        check.setFixedSize(48, 48)
        check.setAlignment(Qt.AlignmentFlag.AlignCenter)
        col.addWidget(check, alignment=Qt.AlignmentFlag.AlignHCenter)
        col.addSpacing(16)

        self._ready_title = QLabel("Đã mở thư mục")
        self._ready_title.setObjectName("welcomeTitle")
        self._ready_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        col.addWidget(self._ready_title)
        col.addSpacing(8)

        self._ready_sub = QLabel("")
        self._ready_sub.setObjectName("welcomeSub")
        self._ready_sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._ready_sub.setWordWrap(True)
        self._ready_sub.setMaximumWidth(360)
        col.addWidget(self._ready_sub, alignment=Qt.AlignmentFlag.AlignHCenter)
        col.addSpacing(18)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        btn_row.addStretch()

        select_all_btn = QPushButton("Chọn tất cả")
        select_all_btn.setProperty("class", "secondaryBtn")
        select_all_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        select_all_btn.clicked.connect(self.select_all_requested.emit)
        btn_row.addWidget(select_all_btn)

        self._start_chat_btn = QPushButton("Bắt đầu chat")
        self._start_chat_btn.setProperty("class", "primaryBtn")
        self._start_chat_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._start_chat_btn.clicked.connect(self.start_chat_requested.emit)
        btn_row.addWidget(self._start_chat_btn)
        btn_row.addStretch()

        col.addLayout(btn_row)
        lay.addLayout(col)
        lay.addStretch(4)
        return page

    # ── API điều khiển từ MainWindow ─────────────────────────────────────────

    def show_empty(self):
        self._stack.setCurrentIndex(_EMPTY)

    def show_scanning(self, folder_name: str):
        self._scan_title.setText(f"Đang lập chỉ mục “{folder_name}”…")
        self._scan_bar.setValue(0)
        self._clear_counts()
        self._stack.setCurrentIndex(_SCANNING)

    def update_progress(self, done: int, total: int, counts: dict):
        pct = int(done * 100 / total) if total else 0
        self._scan_bar.setValue(pct)
        self._clear_counts()
        for ext_type in sorted(counts):
            label, _key = FILE_BADGE_STYLE.get(ext_type, (ext_type.upper(), "other"))
            self._count_row.addWidget(_CountBox(counts[ext_type], label))

    def _clear_counts(self):
        while self._count_row.count():
            item = self._count_row.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

    def show_ready(self, folder_name: str, total_files: int):
        self._ready_title.setText(f"Đã mở thư mục “{folder_name}”")
        self._ready_sub.setText(
            f"Tìm thấy <b>{total_files} file</b> được hỗ trợ. Tick các file bên trái "
            "để đưa vào ngữ cảnh, rồi ra lệnh cho AI.")
        self._stack.setCurrentIndex(_READY)
