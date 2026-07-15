"""Sidebar — segmented Tệp/Thư mục · Trò chuyện mới · lịch sử gần đây · gói dịch vụ."""
from pathlib import Path

from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QFontMetrics, QPainter, QColor, QPen
from PyQt6.QtWidgets import (
    QFrame, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QButtonGroup,
    QScrollArea, QWidget, QProgressBar,
)

from ..constants import EXT_MAP, FILE_BADGE_STYLE

_BADGE_CLASS = {
    "doc": "fileBadgeDoc", "pdf": "fileBadgePdf",
    "xls": "fileBadgeXls", "other": "fileBadgeOther",
}


class RecentItem(QFrame):
    """Một dòng lịch sử — 2 hàng: tiêu đề + (badge loại file · tên file).

    Nền thẻ (trắng khi đang mở / hover) được vẽ trực tiếp bằng QPainter để không
    phụ thuộc vào việc QFrame con có tô background từ stylesheet hay không.
    """

    picked = pyqtSignal(str)

    def __init__(self, path: str, parent=None):
        super().__init__(parent)
        self._path = path
        self._active = False
        self._hover = False
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setToolTip(path)

        p = Path(path)
        ftype = EXT_MAP.get(p.suffix.lower())
        badge_txt, badge_key = FILE_BADGE_STYLE.get(ftype, ("FILE", "other"))

        lay = QVBoxLayout(self)
        lay.setContentsMargins(10, 8, 10, 8)
        lay.setSpacing(3)

        self._title = QLabel(p.stem)
        self._title.setObjectName("recentTitle")
        lay.addWidget(self._title)

        meta = QHBoxLayout()
        meta.setContentsMargins(0, 0, 0, 0)
        meta.setSpacing(5)
        badge = QLabel(badge_txt)
        badge.setProperty("class", _BADGE_CLASS.get(badge_key, "fileBadgeOther"))
        meta.addWidget(badge)
        self._name = QLabel(p.name)
        self._name.setObjectName("recentMeta")
        meta.addWidget(self._name, stretch=1)
        lay.addLayout(meta)

        # Cắt bớt chuỗi dài để không phá bề rộng sidebar cố định.
        self._title.setText(
            QFontMetrics(self._title.font()).elidedText(
                p.stem, Qt.TextElideMode.ElideRight, 168))
        self._name.setText(
            QFontMetrics(self._name.font()).elidedText(
                p.name, Qt.TextElideMode.ElideRight, 128))

    def set_active(self, on: bool):
        self._active = on
        self.update()

    def enterEvent(self, event):
        self._hover = True
        self.update()
        super().enterEvent(event)

    def leaveEvent(self, event):
        self._hover = False
        self.update()
        super().leaveEvent(event)

    def paintEvent(self, event):
        # Vẽ thẻ nền trắng bo góc khi đang mở (hoặc hover) — như nút "Trò chuyện mới".
        if self._active or self._hover:
            painter = QPainter(self)
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            rect = self.rect().adjusted(0, 0, -1, -1)
            painter.setBrush(QColor("#FFFFFF"))
            if self._active:
                painter.setPen(QPen(QColor("#C9C8C0"), 1))
            else:
                painter.setPen(Qt.PenStyle.NoPen)
            painter.drawRoundedRect(rect, 8, 8)
            painter.end()
        super().paintEvent(event)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.picked.emit(self._path)
        super().mousePressEvent(event)


class Sidebar(QFrame):
    new_chat = pyqtSignal()
    file_selected = pyqtSignal(str)   # đường dẫn file
    mode_changed = pyqtSignal(str)    # "file" | "folder"

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("sidebar")
        self.setFixedWidth(230)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(12, 14, 12, 14)
        lay.setSpacing(10)

        # ── Segmented control: Tệp / Thư mục ──────────────────────────────────
        seg = QFrame()
        seg.setObjectName("segControl")
        seg_lay = QHBoxLayout(seg)
        seg_lay.setContentsMargins(3, 3, 3, 3)
        seg_lay.setSpacing(3)

        self._seg_group = QButtonGroup(self)
        self._seg_group.setExclusive(True)
        for key, label in (("file", "Tệp"), ("folder", "Thư mục")):
            btn = QPushButton(label)
            btn.setProperty("class", "segBtn")
            btn.setCheckable(True)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            if key == "file":
                btn.setChecked(True)
            btn.clicked.connect(lambda _, k=key: self.mode_changed.emit(k))
            self._seg_group.addButton(btn)
            seg_lay.addWidget(btn)
        lay.addWidget(seg)

        new_btn = QPushButton("+  Trò chuyện mới")
        new_btn.setObjectName("newFileBtn")
        new_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        new_btn.clicked.connect(self.new_chat.emit)
        lay.addWidget(new_btn)

        recent_lbl = QLabel("Gần đây".upper())
        recent_lbl.setObjectName("recentLabel")
        lay.addWidget(recent_lbl)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setStyleSheet("background: transparent;")
        self._list_host = QWidget()
        self._list_host.setStyleSheet("background: transparent;")
        self._list_lay = QVBoxLayout(self._list_host)
        self._list_lay.setContentsMargins(0, 0, 0, 0)
        self._list_lay.setSpacing(2)
        self._list_lay.addStretch()
        scroll.setWidget(self._list_host)
        lay.addWidget(scroll, stretch=1)

        # ── Thẻ gói dịch vụ + thanh tiến trình ────────────────────────────────
        self._plan_card = QFrame()
        self._plan_card.setObjectName("planCard")
        pc = QVBoxLayout(self._plan_card)
        pc.setContentsMargins(12, 11, 12, 11)
        pc.setSpacing(7)

        head = QHBoxLayout()
        head.setContentsMargins(0, 0, 0, 0)
        self._plan_name = QLabel("Gói Free")
        self._plan_name.setObjectName("planName")
        head.addWidget(self._plan_name)
        head.addStretch()
        self._plan_count = QLabel("")
        self._plan_count.setObjectName("planCount")
        head.addWidget(self._plan_count)
        pc.addLayout(head)

        self._plan_bar = QProgressBar()
        self._plan_bar.setObjectName("planBar")
        self._plan_bar.setTextVisible(False)
        self._plan_bar.setFixedHeight(5)
        self._plan_bar.setMaximum(1)
        self._plan_bar.setValue(0)
        pc.addWidget(self._plan_bar)

        self._plan_hint = QLabel("tác vụ AI còn lại tháng này")
        self._plan_hint.setObjectName("planHint")
        pc.addWidget(self._plan_hint)

        lay.addWidget(self._plan_card)

        self._items: list[RecentItem] = []

    # ── Danh sách gần đây ───────────────────────────────────────────────────────

    def refresh(self, recent_paths: list[str], active_path: str | None = None):
        while self._list_lay.count() > 1:
            item = self._list_lay.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()
        self._items = []

        if not recent_paths:
            empty = QLabel("Chưa có cuộc nào")
            empty.setObjectName("recentEmpty")
            self._list_lay.insertWidget(0, empty)
            return

        for i, path in enumerate(recent_paths):
            it = RecentItem(path)
            it.set_active(bool(active_path) and path == active_path)
            it.picked.connect(self.file_selected.emit)
            self._list_lay.insertWidget(i, it)
            self._items.append(it)

    # ── Gói dịch vụ ─────────────────────────────────────────────────────────────

    def set_plan(self, plan_label: str, remaining, limit):
        self._plan_name.setText(f"Gói {plan_label}")
        if remaining is None or limit is None:
            self._plan_count.setText("")
            self._plan_bar.setMaximum(1)
            self._plan_bar.setValue(0)
        else:
            self._plan_count.setText(f"{remaining}/{limit}")
            self._plan_bar.setMaximum(max(1, int(limit)))
            self._plan_bar.setValue(int(remaining))
