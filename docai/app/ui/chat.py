"""Panel Chat AI — bong bóng hội thoại, file card, quick-action chips, streaming."""
from PySide6.QtCore import Qt, QSize, QTimer, Signal
from PySide6.QtWidgets import (
    QFrame, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QScrollArea, QWidget, QLineEdit,
)

from .icons import paperclip_icon, send_icon


class _Bubble(QLabel):
    def __init__(self, text: str, css_class: str):
        super().__init__(text)
        self.setProperty("class", css_class)
        self.setWordWrap(True)
        self.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.setMaximumWidth(430)


class ChatPanel(QFrame):
    message_sent = Signal(str)
    file_card_clicked = Signal(str)
    attach_requested = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("chatPanel")
        self._stream_bubble: _Bubble | None = None

        lay = QVBoxLayout(self)
        lay.setContentsMargins(1, 1, 1, 1)
        lay.setSpacing(0)

        # ── Header ────────────────────────────────────────────────────────────
        hdr = QFrame()
        hdr.setFixedHeight(46)
        hdr_lay = QHBoxLayout(hdr)
        hdr_lay.setContentsMargins(16, 0, 16, 0)
        title = QLabel("Chat AI")
        title.setObjectName("chatHeader")
        hdr_lay.addWidget(title)
        lay.addWidget(hdr)

        # ── Vùng hội thoại ────────────────────────────────────────────────────
        self._scroll = QScrollArea()
        self._scroll.setObjectName("chatScroll")
        self._scroll.setWidgetResizable(True)
        self._scroll.setFrameShape(QFrame.Shape.NoFrame)
        host = QWidget()
        host.setObjectName("chatMsgHost")
        self._msg_lay = QVBoxLayout(host)
        self._msg_lay.setContentsMargins(16, 8, 16, 8)
        self._msg_lay.setSpacing(10)
        self._msg_lay.addStretch()
        self._scroll.setWidget(host)
        lay.addWidget(self._scroll, stretch=1)

        # ── Quick-action chips ────────────────────────────────────────────────
        self._chips_host = QWidget()
        self._chips_host.setObjectName("chatChipsHost")
        self._chips_lay = QHBoxLayout(self._chips_host)
        self._chips_lay.setContentsMargins(16, 4, 16, 8)
        self._chips_lay.setSpacing(8)
        self._chips_lay.addStretch()
        lay.addWidget(self._chips_host)

        # ── Ô nhập ────────────────────────────────────────────────────────────
        input_area = QFrame()
        input_area.setObjectName("inputArea")
        in_lay = QHBoxLayout(input_area)
        in_lay.setContentsMargins(10, 10, 10, 12)
        in_lay.setSpacing(6)

        self.attach_btn = QPushButton()
        self.attach_btn.setObjectName("attachBtn")
        self.attach_btn.setIcon(paperclip_icon())
        self.attach_btn.setIconSize(QSize(16, 16))
        self.attach_btn.setFixedSize(30, 30)
        self.attach_btn.setToolTip("Đính kèm file")
        self.attach_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.attach_btn.clicked.connect(self.attach_requested.emit)
        in_lay.addWidget(self.attach_btn)

        self.input_box = QLineEdit()
        self.input_box.setObjectName("inputBox")
        self.input_box.setPlaceholderText("Nhập yêu cầu bằng tiếng Việt…")
        self.input_box.returnPressed.connect(self._submit)
        in_lay.addWidget(self.input_box, stretch=1)

        self.send_btn = QPushButton()
        self.send_btn.setObjectName("sendBtn")
        self.send_btn.setIcon(send_icon())
        self.send_btn.setIconSize(QSize(15, 15))
        self.send_btn.setFixedSize(32, 32)
        self.send_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.send_btn.clicked.connect(self._submit)
        in_lay.addWidget(self.send_btn)
        lay.addWidget(input_area)

    # ── Gửi tin ───────────────────────────────────────────────────────────────

    def _submit(self):
        text = self.input_box.text().strip()
        if not text:
            return
        self.input_box.clear()
        self.message_sent.emit(text)

    def send_text(self, text: str):
        """Gửi hộ (chip click)."""
        self.message_sent.emit(text)

    def set_enabled(self, on: bool):
        self.input_box.setEnabled(on)
        self.send_btn.setEnabled(on)

    def clear(self):
        """Xóa toàn bộ hội thoại — dùng khi bắt đầu trò chuyện mới."""
        self._stream_bubble = None
        while self._msg_lay.count() > 1:
            item = self._msg_lay.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
            elif item.layout():
                lay = item.layout()
                while lay.count():
                    sub = lay.takeAt(0)
                    if sub.widget():
                        sub.widget().deleteLater()
                lay.deleteLater()
        self.input_box.clear()

    # ── Chips ─────────────────────────────────────────────────────────────────

    def set_chips(self, labels: list[str]):
        while self._chips_lay.count() > 1:
            item = self._chips_lay.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        for chip_index, label in enumerate(labels):
            chip = QPushButton(label)
            chip.setProperty("class", "chip")
            chip.setCursor(Qt.CursorShape.PointingHandCursor)
            chip.clicked.connect(lambda _, chip_label=label: self.send_text(chip_label))
            self._chips_lay.insertWidget(chip_index, chip)

    # ── Thêm bong bóng ────────────────────────────────────────────────────────

    def _add_row(self, widget, align_right: bool):
        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        if align_right:
            row.addStretch()
            row.addWidget(widget)
        else:
            row.addWidget(widget)
            row.addStretch()
        self._msg_lay.insertLayout(self._msg_lay.count() - 1, row)
        self._scroll_to_bottom()

    def add_widget_row(self, widget, full_width: bool = True):
        """Chèn 1 widget tuỳ ý vào luồng hội thoại (thẻ chọn loại file, thẻ
        cảnh báo…) — dùng cho các bước của luồng "Tạo tài liệu mới bằng chat"."""
        if full_width:
            self._msg_lay.insertWidget(self._msg_lay.count() - 1, widget)
            self._scroll_to_bottom()
        else:
            self._add_row(widget, align_right=False)

    def add_user(self, text: str):
        self._add_row(_Bubble(text, "bubbleUser"), align_right=True)

    def add_ai(self, text: str):
        """Bong bóng AI tĩnh (không streaming) — dùng cho các câu dẫn ngắn
        trước thẻ chọn loại file / thẻ gợi ý, khi không cần hiệu ứng gõ chữ."""
        self._add_row(_Bubble(text, "bubbleAI"), align_right=False)

    def add_system(self, text: str):
        lbl = QLabel(text)
        lbl.setProperty("class", "bubbleSystem")
        lbl.setWordWrap(True)
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._msg_lay.insertWidget(self._msg_lay.count() - 1, lbl)
        self._scroll_to_bottom()

    def add_file_card(self, file_name: str, badge: str = "XLS", note: str = "",
                      full_path: str | None = None):
        """`full_path` (tùy chọn): đường dẫn thật để mở khi bấm vào thẻ — nhãn
        hiển thị vẫn dùng `file_name` (ngắn gọn). Bỏ trống thì click phát ra
        chính `file_name` như trước (các nơi gọi cũ, artifact còn là mock)."""
        label = f"[{badge}]  {file_name}" + (f"  ·  {note}" if note else "")
        card = QPushButton(label)
        card.setProperty("class", "fileCard")
        card.setCursor(Qt.CursorShape.PointingHandCursor)
        card.clicked.connect(lambda: self.file_card_clicked.emit(full_path or file_name))
        self._add_row(card, align_right=False)

    # ── Streaming AI ──────────────────────────────────────────────────────────

    def start_ai(self):
        self._stream_bubble = _Bubble("…", "bubbleAI")
        self._add_row(self._stream_bubble, align_right=False)

    def stream_ai(self, text_so_far: str):
        if self._stream_bubble:
            self._stream_bubble.setText(text_so_far)
            self._scroll_to_bottom()

    def end_ai(self):
        self._stream_bubble = None

    def _scroll_to_bottom(self):
        QTimer.singleShot(0, lambda: self._scroll.verticalScrollBar().setValue(
            self._scroll.verticalScrollBar().maximum()))
