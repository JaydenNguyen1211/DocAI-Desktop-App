"""Chat AI trung tâm — màn hình mở app: chat sẵn sàng ngay, file là ngữ cảnh tùy chọn."""
from PySide6.QtCore import Qt, QSize, Signal
from PySide6.QtWidgets import (
    QFrame, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QLineEdit, QFileDialog,
)

from ..constants import EMPTY_CHIPS, FILE_DIALOG_FILTER, EXT_MAP
from .icons import paperclip_icon, send_icon, diamond_icon

from ...logging_config import get_logger, log_call
from ...strings import Common, EmptyState as S

logger = get_logger(__name__)


class CentralChat(QFrame):
    """Trạng thái 1 — chưa có file. Ô chat trung tâm + chip gợi ý + đính kèm."""

    message_sent = Signal(str)   # gõ lệnh / bấm chip → bắt đầu trò chuyện
    file_chosen = Signal(str)    # đính kèm / kéo-thả file → vào workspace

    @log_call
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("centralChat")
        self.setAcceptDrops(True)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(48, 40, 48, 40)
        outer.addStretch(3)

        # ── Cụm trung tâm (logo · tiêu đề · chip · ô nhập) ─────────────────────
        col = QVBoxLayout()
        col.setSpacing(0)

        logo = QLabel("D")
        logo.setObjectName("welcomeLogo")
        logo.setFixedSize(52, 52)
        logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        col.addWidget(logo, alignment=Qt.AlignmentFlag.AlignHCenter)
        col.addSpacing(22)

        title = QLabel(S.TITLE)
        title.setObjectName("welcomeTitle")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        col.addWidget(title)
        col.addSpacing(6)

        sub = QLabel(S.SUBTITLE)
        sub.setObjectName("welcomeSub")
        sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        col.addWidget(sub)
        col.addSpacing(24)

        # Chip gợi ý
        chips_wrap = QVBoxLayout()
        chips_wrap.setSpacing(8)
        row = None
        for chip_index, (label, color) in enumerate(EMPTY_CHIPS):
            if chip_index % 2 == 0:
                row = QHBoxLayout()
                row.setSpacing(8)
                row.addStretch()
                chips_wrap.addLayout(row)
            chip = QPushButton(label)
            chip.setProperty("class", "chip")
            chip.setIcon(diamond_icon(color))
            chip.setIconSize(QSize(10, 10))
            chip.setCursor(Qt.CursorShape.PointingHandCursor)
            chip.clicked.connect(lambda _, chip_label=label: self.message_sent.emit(chip_label))
            row.addWidget(chip)
            if chip_index % 2 == 1 or chip_index == len(EMPTY_CHIPS) - 1:
                row.addStretch()
        col.addLayout(chips_wrap)
        col.addSpacing(18)

        # ── Ô nhập trung tâm (2 hàng: text ở trên · công cụ ở dưới) ────────────
        self._input_card = QFrame()
        self._input_card.setObjectName("welcomeInput")
        self._input_card.setFixedWidth(460)
        card_lay = QVBoxLayout(self._input_card)
        card_lay.setContentsMargins(14, 11, 8, 8)
        card_lay.setSpacing(8)

        self.input_box = QLineEdit()
        self.input_box.setObjectName("welcomeInputBox")
        self.input_box.setPlaceholderText(Common.INPUT_PLACEHOLDER)
        self.input_box.returnPressed.connect(self._submit)
        card_lay.addWidget(self.input_box)

        tools = QHBoxLayout()
        tools.setContentsMargins(0, 0, 0, 0)
        tools.setSpacing(0)

        attach = QPushButton()
        attach.setObjectName("attachBtn")
        attach.setIcon(paperclip_icon())
        attach.setIconSize(QSize(17, 17))
        attach.setFixedSize(30, 30)
        attach.setToolTip(S.ATTACH_TOOLTIP)
        attach.setCursor(Qt.CursorShape.PointingHandCursor)
        attach.clicked.connect(self._pick_file)
        tools.addWidget(attach)
        tools.addStretch()

        send = QPushButton()
        send.setObjectName("welcomeSendBtn")
        send.setIcon(send_icon())
        send.setIconSize(QSize(15, 15))
        send.setFixedSize(34, 34)
        send.setCursor(Qt.CursorShape.PointingHandCursor)
        send.clicked.connect(self._submit)
        tools.addWidget(send)

        card_lay.addLayout(tools)

        col.addWidget(self._input_card, alignment=Qt.AlignmentFlag.AlignHCenter)

        outer.addLayout(col)
        outer.addStretch(4)

    # ── Gửi / đính kèm ─────────────────────────────────────────────────────────

    @log_call
    def _submit(self):
        text = self.input_box.text().strip()
        if not text:
            return
        self.input_box.clear()
        self.message_sent.emit(text)

    @log_call
    def _pick_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, S.PICK_FILE_TITLE, "", FILE_DIALOG_FILTER)
        if path:
            self.file_chosen.emit(path)

    # ── Kéo-thả file vào màn hình chat ─────────────────────────────────────────

    @log_call
    def _set_drag_over(self, on: bool):
        self._input_card.setProperty("dragOver", "true" if on else "false")
        self._input_card.style().unpolish(self._input_card)
        self._input_card.style().polish(self._input_card)

    @log_call
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            path = event.mimeData().urls()[0].toLocalFile()
            if "." + path.rsplit(".", 1)[-1].lower() in EXT_MAP:
                self._set_drag_over(True)
                event.acceptProposedAction()
                return
        event.ignore()

    @log_call
    def dragLeaveEvent(self, event):
        self._set_drag_over(False)

    @log_call
    def dropEvent(self, event):
        self._set_drag_over(False)
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if "." + path.rsplit(".", 1)[-1].lower() in EXT_MAP:
                self.file_chosen.emit(path)
                return
