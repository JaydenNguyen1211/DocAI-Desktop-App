"""Cây file của thư mục làm việc (chế độ Thư mục) — lọc theo loại, nhóm theo
thư mục con cấp 1. Dùng trong Sidebar, ở trạng thái "sẵn sàng".

Chỉ là DANH SÁCH THAM CHIẾU cho Chat Panel (luôn mở độc lập, không cần thao
tác gì ở đây trước) — không có khái niệm tick chọn; file nào là input do
người dùng nói rõ trong chat, xem `modules.common.doc_set.resolve_input_files`."""
from PySide6.QtCore import Qt, QSize, Signal
from PySide6.QtGui import QFontMetrics
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QButtonGroup, QFrame, QScrollArea,
)

_GROUP_ROW_H = 24
_FILE_ROW_H = 24

from ..constants import FILE_BADGE_STYLE
from .icons import folder_icon_filled, chevron_icon
from .folder_scan import ScannedFile

from ...logging_config import get_logger, log_call
from ...strings import Common, FolderTree as S

logger = get_logger(__name__)

_BADGE_CLASS = {
    "doc": "fileBadgeDoc", "pdf": "fileBadgePdf",
    "xls": "fileBadgeXls", "other": "fileBadgeOther",
}


class _GroupRow(QFrame):
    """Dòng thư mục con — mở/thu gọn danh sách file bên trong."""

    toggled = Signal()

    @log_call
    def __init__(self, name: str, count: int, parent=None):
        super().__init__(parent)
        self._expanded = True
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setFixedHeight(_GROUP_ROW_H)

        lay = QHBoxLayout(self)
        lay.setContentsMargins(3, 0, 3, 0)
        lay.setSpacing(5)

        self._chevron = QLabel()
        self._chevron.setPixmap(chevron_icon(True).pixmap(QSize(9, 9)))
        lay.addWidget(self._chevron)

        folder_lbl = QLabel()
        folder_lbl.setPixmap(folder_icon_filled().pixmap(QSize(13, 13)))
        lay.addWidget(folder_lbl)

        name_lbl = QLabel(name)
        name_lbl.setObjectName("folderGroupName")
        lay.addWidget(name_lbl)

        self._count_lbl = QLabel(f"({count})")
        self._count_lbl.setObjectName("folderGroupCount")
        lay.addWidget(self._count_lbl)
        lay.addStretch()

    @log_call
    def set_count(self, count: int):
        self._count_lbl.setText(f"({count})")

    @log_call
    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._expanded = not self._expanded
            self._chevron.setPixmap(chevron_icon(self._expanded).pixmap(QSize(9, 9)))
            self.toggled.emit()
        super().mousePressEvent(event)

    @log_call
    def is_expanded(self) -> bool:
        return self._expanded


class _FileRow(QFrame):
    """Dòng file — badge loại · tên file · xóa khỏi danh sách."""

    remove_clicked = Signal()

    @log_call
    def __init__(self, sf: ScannedFile, parent=None):
        super().__init__(parent)
        self.file = sf
        self.setFixedHeight(_FILE_ROW_H)

        lay = QHBoxLayout(self)
        lay.setContentsMargins(20, 0, 3, 0)
        lay.setSpacing(6)

        label, key = FILE_BADGE_STYLE.get(sf.ext_type, ("FILE", "other"))
        badge = QLabel(label)
        badge.setProperty("class", _BADGE_CLASS.get(key, "fileBadgeOther") + "Mini")
        lay.addWidget(badge)

        name_lbl = QLabel()
        name_lbl.setObjectName("folderFileName")
        name_lbl.setText(
            QFontMetrics(name_lbl.font()).elidedText(sf.name, Qt.TextElideMode.ElideMiddle, 130))
        name_lbl.setToolTip(sf.name)
        lay.addWidget(name_lbl, stretch=1)

        self._delete_btn = QPushButton("✕")
        self._delete_btn.setObjectName("folderFileDeleteBtn")
        self._delete_btn.setFixedSize(14, 14)
        self._delete_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._delete_btn.setToolTip(Common.REMOVE_FROM_LIST)
        self._delete_btn.clicked.connect(self.remove_clicked.emit)
        lay.addWidget(self._delete_btn)


class FolderTree(QWidget):
    """Danh sách file quét được — lọc theo loại, chỉ để xem/tham chiếu."""

    file_removed = Signal(str)         # đường dẫn file vừa bị xóa khỏi danh sách

    @log_call
    def __init__(self, parent=None):
        super().__init__(parent)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(8)

        self._chips_row = QHBoxLayout()
        self._chips_row.setSpacing(5)
        outer.addLayout(self._chips_row)
        self._chip_group = QButtonGroup(self)
        self._chip_group.setExclusive(True)
        self._chips: dict[str | None, QPushButton] = {}

        scroll = QScrollArea()
        scroll.setObjectName("folderTreeScroll")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._list_host = QWidget()
        self._list_host.setObjectName("folderTreeListHost")
        self._list_lay = QVBoxLayout(self._list_host)
        self._list_lay.setContentsMargins(0, 0, 0, 0)
        self._list_lay.setSpacing(1)
        self._list_lay.addStretch()
        scroll.setWidget(self._list_host)
        outer.addWidget(scroll, stretch=1)

        self._rows: list[_FileRow] = []
        self._groups: dict[str, _GroupRow] = {}
        self._group_rows: dict[str, list[_FileRow]] = {}
        self._active_filter: str | None = None

    # ── Nạp danh sách file mới ─────────────────────────────────────────────

    @log_call
    def set_files(self, files: list[ScannedFile], counts: dict[str, int]):
        self._clear()

        # Chip lọc: "Tất cả" + 1 chip cho mỗi loại tìm thấy.
        all_chip = self._make_chip(S.FILTER_ALL, None)
        all_chip.setChecked(True)
        self._chips_row.addWidget(all_chip)
        for ext_type in sorted(counts):
            label, _key = FILE_BADGE_STYLE.get(ext_type, (ext_type.upper(), "other"))
            self._chips_row.addWidget(self._make_chip(label, ext_type))
        self._chips_row.addStretch()
        self._active_filter = None

        # Nhóm theo thư mục con cấp 1 — giữ thứ tự xuất hiện đầu tiên.
        order: list[str] = []
        by_group: dict[str, list[ScannedFile]] = {}
        for sf in files:
            if sf.group not in by_group:
                by_group[sf.group] = []
                order.append(sf.group)
            by_group[sf.group].append(sf)

        for group in order:
            items = by_group[group]
            insert_at = self._list_lay.count() - 1   # trước stretch cuối danh sách
            if group:
                header = _GroupRow(group, len(items))
                header.toggled.connect(lambda group_name=group: self._apply_visibility(group_name))
                self._list_lay.insertWidget(insert_at, header)
                insert_at += 1
                self._groups[group] = header
            rows = []
            for sf in items:
                row = _FileRow(sf)
                row.remove_clicked.connect(lambda row_widget=row: self._remove_row(row_widget))
                self._list_lay.insertWidget(insert_at, row)
                insert_at += 1
                self._rows.append(row)
                rows.append(row)
            self._group_rows[group] = rows
            self._apply_visibility(group)

    @log_call
    def _clear(self):
        while self._list_lay.count() > 1:
            item = self._list_lay.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        while self._chips_row.count():
            item = self._chips_row.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        for btn in list(self._chip_group.buttons()):
            self._chip_group.removeButton(btn)
        self._rows = []
        self._groups = {}
        self._group_rows = {}
        self._chips = {}

    # ── Chip lọc theo loại ──────────────────────────────────────────────────

    @log_call
    def _make_chip(self, label: str, ext_type: str | None) -> QPushButton:
        chip = QPushButton(label)
        chip.setProperty("class", "filterChip")
        chip.setCheckable(True)
        chip.setFixedHeight(22)   # bằng 2× border-radius — bo tròn đủ vòng như thiết kế
        chip.setCursor(Qt.CursorShape.PointingHandCursor)
        chip.clicked.connect(lambda _, chip_ext_type=ext_type: self._set_filter(chip_ext_type))
        self._chip_group.addButton(chip)
        self._chips[ext_type] = chip
        return chip

    @log_call
    def _set_filter(self, ext_type: str | None):
        self._active_filter = ext_type
        for group in self._group_rows:
            self._apply_visibility(group)

    @log_call
    def _apply_visibility(self, group: str):
        header = self._groups.get(group)
        expanded = header.is_expanded() if header else True
        any_visible = False
        for row in self._group_rows.get(group, []):
            matches = self._active_filter is None or row.file.ext_type == self._active_filter
            visible = matches and expanded
            row.setVisible(visible)
            any_visible = any_visible or matches
        if header is not None:
            header.setVisible(any_visible)

    # ── Xóa 1 file khỏi danh sách ─────────────────────────────────────────────

    @log_call
    def _remove_row(self, row: "_FileRow"):
        group = row.file.group
        path = row.file.path

        self._list_lay.removeWidget(row)
        row.deleteLater()
        if row in self._rows:
            self._rows.remove(row)
        rows = self._group_rows.get(group, [])
        if row in rows:
            rows.remove(row)

        header = self._groups.get(group)
        if not rows:
            if header is not None:
                self._list_lay.removeWidget(header)
                header.deleteLater()
                del self._groups[group]
            self._group_rows.pop(group, None)
        elif header is not None:
            header.set_count(len(rows))

        self.file_removed.emit(path)
