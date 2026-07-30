"""Sidebar — segmented Tệp/Thư mục · Trò chuyện mới · lịch sử gần đây · gói dịch vụ.

Chế độ Thư mục có 3 trạng thái con (rỗng · đang quét · sẵn sàng), tự thay đổi
qua các phương thức `show_folder_*` do MainWindow gọi khi mở/quét thư mục."""
from pathlib import Path

from PySide6.QtCore import Qt, QSize, Signal
from PySide6.QtGui import QFontMetrics, QPainter, QColor, QPen
from PySide6.QtWidgets import (
    QFrame, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QButtonGroup,
    QScrollArea, QWidget, QProgressBar, QStackedWidget,
)

from ..constants import EXT_MAP, FILE_BADGE_STYLE
from .icons import folder_icon
from .folder_tree import FolderTree
from .folder_scan import ScannedFile

_MODE_FILE, _MODE_FOLDER = 0, 1
_FOLDER_SCANNING, _FOLDER_READY = 1, 2

_BADGE_CLASS = {
    "doc": "fileBadgeDoc", "pdf": "fileBadgePdf",
    "xls": "fileBadgeXls", "other": "fileBadgeOther",
}


class RecentItem(QFrame):
    """Một dòng lịch sử — 2 hàng: tiêu đề + (badge loại file · tên file).

    Nền thẻ (trắng khi đang mở / hover) được vẽ trực tiếp bằng QPainter để không
    phụ thuộc vào việc QFrame con có tô background từ stylesheet hay không.
    """

    picked = Signal(str)
    remove_clicked = Signal(str)

    def __init__(self, path: str, parent=None):
        super().__init__(parent)
        self._path = path
        self._active = False
        self._hover = False
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setToolTip(path)

        path_obj = Path(path)
        ftype = EXT_MAP.get(path_obj.suffix.lower())
        badge_txt, badge_key = FILE_BADGE_STYLE.get(ftype, ("FILE", "other"))

        lay = QVBoxLayout(self)
        lay.setContentsMargins(10, 8, 10, 8)
        lay.setSpacing(3)

        title_row = QHBoxLayout()
        title_row.setContentsMargins(0, 0, 0, 0)
        title_row.setSpacing(4)
        self._title = QLabel(path_obj.stem)
        self._title.setObjectName("recentTitle")
        title_row.addWidget(self._title, stretch=1)

        self._delete_btn = QPushButton("✕")
        self._delete_btn.setObjectName("recentDeleteBtn")
        self._delete_btn.setFixedSize(16, 16)
        self._delete_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._delete_btn.setToolTip("Xóa khỏi danh sách")
        self._delete_btn.setVisible(False)
        self._delete_btn.clicked.connect(lambda: self.remove_clicked.emit(self._path))
        title_row.addWidget(self._delete_btn)
        lay.addLayout(title_row)

        meta = QHBoxLayout()
        meta.setContentsMargins(0, 0, 0, 0)
        meta.setSpacing(5)
        badge = QLabel(badge_txt)
        badge.setProperty("class", _BADGE_CLASS.get(badge_key, "fileBadgeOther"))
        meta.addWidget(badge)
        self._name = QLabel(path_obj.name)
        self._name.setObjectName("recentMeta")
        meta.addWidget(self._name, stretch=1)
        lay.addLayout(meta)

        # Cắt bớt chuỗi dài để không phá bề rộng sidebar cố định.
        self._title.setText(
            QFontMetrics(self._title.font()).elidedText(
                path_obj.stem, Qt.TextElideMode.ElideRight, 168))
        self._name.setText(
            QFontMetrics(self._name.font()).elidedText(
                path_obj.name, Qt.TextElideMode.ElideRight, 128))

    def set_active(self, on: bool):
        self._active = on
        self.update()

    def enterEvent(self, event):
        self._hover = True
        self._delete_btn.setVisible(True)
        self.update()
        super().enterEvent(event)

    def leaveEvent(self, event):
        self._hover = False
        self._delete_btn.setVisible(False)
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
    new_chat = Signal()
    file_selected = Signal(str)   # đường dẫn file
    mode_changed = Signal(str)    # "file" | "folder"
    recent_deleted = Signal(str)  # bấm xóa 1 item trong danh sách "Gần đây"

    open_folder = Signal()             # bấm "Mở thư mục làm việc" (trống)
    change_folder = Signal()           # bấm "Đổi" (đang quét / sẵn sàng)
    folder_selection_changed = Signal(list)   # list[str] file đã tick
    folder_file_removed = Signal(str)         # bấm xóa 1 file trong cây thư mục

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
            btn.clicked.connect(lambda _, seg_key=key: self._on_seg_clicked(seg_key))
            self._seg_group.addButton(btn)
            seg_lay.addWidget(btn)
        lay.addWidget(seg)

        # ── Thân sidebar: nội dung Tệp / Thư mục thay nhau ─────────────────────
        self._body_stack = QStackedWidget()
        self._body_stack.addWidget(self._build_file_body())
        self._body_stack.addWidget(self._build_folder_body())
        lay.addWidget(self._body_stack, stretch=1)

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

    # ── Chuyển đổi Tệp ⇄ Thư mục ──────────────────────────────────────────────

    def _on_seg_clicked(self, mode: str):
        self._body_stack.setCurrentIndex(_MODE_FOLDER if mode == "folder" else _MODE_FILE)
        self.mode_changed.emit(mode)

    # ── Thân "Tệp": Trò chuyện mới + lịch sử gần đây ─────────────────────────

    def _build_file_body(self) -> QWidget:
        page = QWidget()
        lay = QVBoxLayout(page)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(10)

        new_btn = QPushButton("+  Trò chuyện mới")
        new_btn.setObjectName("newFileBtn")
        new_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        new_btn.clicked.connect(self.new_chat.emit)
        lay.addWidget(new_btn)

        recent_lbl = QLabel("Gần đây".upper())
        recent_lbl.setObjectName("recentLabel")
        lay.addWidget(recent_lbl)

        scroll = QScrollArea()
        scroll.setObjectName("recentScroll")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._list_host = QWidget()
        self._list_host.setObjectName("recentListHost")
        self._list_lay = QVBoxLayout(self._list_host)
        self._list_lay.setContentsMargins(0, 0, 0, 0)
        self._list_lay.setSpacing(2)
        self._list_lay.addStretch()
        scroll.setWidget(self._list_host)
        lay.addWidget(scroll, stretch=1)
        return page

    # ── Thân "Thư mục": trống → đang quét → sẵn sàng ─────────────────────────

    def _build_folder_body(self) -> QWidget:
        page = QWidget()
        lay = QVBoxLayout(page)
        lay.setContentsMargins(0, 0, 0, 0)

        self._folder_stack = QStackedWidget()
        lay.addWidget(self._folder_stack)

        self._folder_stack.addWidget(self._build_folder_empty())
        self._folder_stack.addWidget(self._build_folder_scanning())
        self._folder_stack.addWidget(self._build_folder_ready())
        return page

    def _build_folder_empty(self) -> QWidget:
        page = QWidget()
        lay = QVBoxLayout(page)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(10)

        btn = QPushButton("  Mở thư mục làm việc")
        btn.setObjectName("openFolderBtn")
        btn.setIcon(folder_icon("#FFFFFF"))
        btn.setIconSize(QSize(15, 15))
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.clicked.connect(self.open_folder.emit)
        lay.addWidget(btn)

        lay.addSpacing(6)
        icon = QLabel()
        icon.setObjectName("folderEmptySmallIcon")
        icon.setFixedSize(44, 44)
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon.setPixmap(folder_icon("#B4B0A8", 20).pixmap(QSize(20, 20)))
        lay.addWidget(icon, alignment=Qt.AlignmentFlag.AlignHCenter)

        hint = QLabel(
            "Chưa mở thư mục nào.\nMở một thư mục để tick nhiều file làm ngữ cảnh.")
        hint.setObjectName("folderEmptyHint")
        hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        hint.setWordWrap(True)
        lay.addWidget(hint)
        lay.addStretch()
        return page

    def _build_folder_scanning(self) -> QWidget:
        page = QWidget()
        lay = QVBoxLayout(page)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(10)

        self._scan_badge, badge_lbl = self._make_path_badge(change_link=False)
        self._scan_path_lbl = badge_lbl
        lay.addWidget(self._scan_badge)

        lay.addStretch(1)
        spin = QProgressBar()
        spin.setObjectName("folderSpinBar")
        spin.setFixedHeight(4)
        spin.setTextVisible(False)
        spin.setMaximum(0)   # chế độ "đang chạy" — vô định
        lay.addWidget(spin)

        self._scan_status_lbl = QLabel("Đang quét thư mục…")
        self._scan_status_lbl.setObjectName("folderScanStatus")
        self._scan_status_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._scan_status_lbl.setWordWrap(True)
        lay.addWidget(self._scan_status_lbl)
        lay.addStretch(2)
        return page

    def _build_folder_ready(self) -> QWidget:
        page = QWidget()
        lay = QVBoxLayout(page)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(9)

        self._ready_badge, badge_lbl = self._make_path_badge(change_link=True)
        self._ready_path_lbl = badge_lbl
        lay.addWidget(self._ready_badge)

        self._folder_tree = FolderTree()
        self._folder_tree.selection_changed.connect(self._on_folder_tree_selection)
        self._folder_tree.file_removed.connect(self.folder_file_removed.emit)
        lay.addWidget(self._folder_tree, stretch=1)

        footer = QFrame()
        footer.setObjectName("folderFooter")
        footer_lay = QHBoxLayout(footer)
        footer_lay.setContentsMargins(10, 8, 10, 8)
        self._folder_footer_lbl = QLabel("Chưa chọn file")
        self._folder_footer_lbl.setObjectName("folderFooterLabel")
        footer_lay.addWidget(self._folder_footer_lbl)
        footer_lay.addStretch()
        select_all_lbl = QPushButton("Chọn tất cả")
        select_all_lbl.setObjectName("folderFooterAction")
        select_all_lbl.setCursor(Qt.CursorShape.PointingHandCursor)
        select_all_lbl.clicked.connect(lambda: self._folder_tree.select_all())
        footer_lay.addWidget(select_all_lbl)
        lay.addWidget(footer)
        return page

    def _make_path_badge(self, change_link: bool) -> tuple[QFrame, QLabel]:
        badge = QFrame()
        badge.setObjectName("folderPathBadge")
        bl = QHBoxLayout(badge)
        bl.setContentsMargins(9, 7, 9, 7)
        bl.setSpacing(6)

        icon = QLabel()
        icon.setPixmap(folder_icon("#2A6FDB", 13).pixmap(QSize(13, 13)))
        bl.addWidget(icon)

        path_lbl = QLabel("")
        path_lbl.setObjectName("folderPathLabel")
        bl.addWidget(path_lbl, stretch=1)

        if change_link:
            change_btn = QPushButton("Đổi")
            change_btn.setObjectName("folderChangeBtn")
            change_btn.setCursor(Qt.CursorShape.PointingHandCursor)
            change_btn.clicked.connect(self.change_folder.emit)
            bl.addWidget(change_btn)

        return badge, path_lbl

    # ── Điều khiển từ MainWindow khi mở / quét thư mục ───────────────────────

    def show_folder_scanning(self, folder_path: str):
        self._body_stack.setCurrentIndex(_MODE_FOLDER)
        self._folder_stack.setCurrentIndex(_FOLDER_SCANNING)
        self._scan_path_lbl.setText(
            QFontMetrics(self._scan_path_lbl.font()).elidedText(
                folder_path, Qt.TextElideMode.ElideMiddle, 140))
        self._scan_status_lbl.setText("Đang quét thư mục…")

    def update_folder_scan_progress(self, done: int, total: int):
        suffix = f" / {total}" if total else ""
        self._scan_status_lbl.setText(f"Đang quét thư mục…\n{done}{suffix} file")

    def show_folder_ready(self, folder_name: str, files: list[ScannedFile], counts: dict):
        self._folder_stack.setCurrentIndex(_FOLDER_READY)
        self._ready_path_lbl.setText(
            QFontMetrics(self._ready_path_lbl.font()).elidedText(
                folder_name, Qt.TextElideMode.ElideMiddle, 110))
        self._folder_tree.set_files(files, counts)
        self._update_folder_footer([])

    def select_all_folder_files(self):
        self._folder_tree.select_all()

    def _on_folder_tree_selection(self, paths: list[str]):
        self._update_folder_footer(paths)
        self.folder_selection_changed.emit(paths)

    def _update_folder_footer(self, paths: list[str]):
        if paths:
            self._folder_footer_lbl.setText(f"{len(paths)} file đã chọn")
        else:
            self._folder_footer_lbl.setText("Chưa chọn file")

    # ── Danh sách gần đây ───────────────────────────────────────────────────────

    def refresh(self, recent_paths: list[str], active_path: str | None = None):
        while self._list_lay.count() > 1:
            item = self._list_lay.takeAt(0)
            item_widget = item.widget()
            if item_widget is not None:
                item_widget.deleteLater()
        self._items = []

        if not recent_paths:
            empty = QLabel("Chưa có cuộc nào")
            empty.setObjectName("recentEmpty")
            self._list_lay.insertWidget(0, empty)
            return

        for item_index, path in enumerate(recent_paths):
            it = RecentItem(path)
            it.set_active(bool(active_path) and path == active_path)
            it.picked.connect(self.file_selected.emit)
            it.remove_clicked.connect(self.recent_deleted.emit)
            self._list_lay.insertWidget(item_index, it)
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
