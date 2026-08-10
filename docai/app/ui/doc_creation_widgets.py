"""Thẻ hội thoại nội tuyến cho luồng "Tạo tài liệu mới bằng chat" (7.2, EC4,
EC5, EC6) — chèn vào `ChatPanel` qua `add_widget_row()`."""
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFrame, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QButtonGroup,
)

from ...logging_config import get_logger, log_call
from ...strings import DocCreation as S

logger = get_logger(__name__)


class _TypeCard(QPushButton):
    @log_call
    def __init__(self, key: str, label: str, badge: str, fg: str, bg: str):
        super().__init__()
        self.key = key
        self.setCheckable(True)
        self.setProperty("class", "docTypeCard")
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setFixedSize(96, 76)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(6, 10, 6, 8)
        lay.setSpacing(7)

        icon = QLabel(badge)
        icon.setFixedSize(34, 34)
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon.setStyleSheet(
            f"background:{bg}; color:{fg}; border-radius:8px; "
            "font-family:'Consolas','IBM Plex Mono',monospace; "
            "font-size:10px; font-weight:600;"
        )
        lay.addWidget(icon, alignment=Qt.AlignmentFlag.AlignHCenter)

        text = QLabel(label)
        text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        text.setStyleSheet("font-size:11px; font-weight:600; background:transparent;")
        lay.addWidget(text, alignment=Qt.AlignmentFlag.AlignHCenter)


class DocTypePicker(QFrame):
    """7.2 — thẻ chọn Word/Excel/PowerPoint + nút xác nhận."""

    confirmed = Signal(str)   # file_type đã chọn

    @log_call
    def __init__(self, options: list[tuple[str, str, str, str, str]],
                 default_key: str | None, page_count: int, parent=None):
        super().__init__(parent)
        self.setObjectName("docTypePicker")
        self._page_count = page_count
        self._labels = {key: label for key, label, *_ in options}

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 2, 0, 8)
        outer.setSpacing(10)

        cards_row = QHBoxLayout()
        cards_row.setSpacing(10)
        self._group = QButtonGroup(self)
        self._group.setExclusive(True)
        self._cards: dict[str, _TypeCard] = {}
        for key, label, badge, fg, bg in options:
            card = _TypeCard(key, label, badge, fg, bg)
            card.toggled.connect(self._update_confirm_label)
            self._group.addButton(card)
            self._cards[key] = card
            cards_row.addWidget(card)
        cards_row.addStretch()
        outer.addLayout(cards_row)

        self.confirm_btn = QPushButton()
        self.confirm_btn.setObjectName("createConfirmBtn")
        self.confirm_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.confirm_btn.clicked.connect(self._on_confirm)
        outer.addWidget(self.confirm_btn)

        selected = default_key if default_key in self._cards else next(iter(self._cards))
        self._cards[selected].setChecked(True)
        self._update_confirm_label()

    @log_call
    def _selected_key(self) -> str:
        for key, card in self._cards.items():
            if card.isChecked():
                return key
        return next(iter(self._cards))

    @log_call
    def _update_confirm_label(self, *_args):
        key = self._selected_key()
        label = self._labels[key]
        if key != "excel" and self._page_count:
            self.confirm_btn.setText(S.CONFIRM_WITH_PAGES.format(label=label, page_count=self._page_count))
        else:
            self.confirm_btn.setText(S.CONFIRM.format(label=label))

    @log_call
    def _on_confirm(self):
        key = self._selected_key()
        for card in self._cards.values():
            card.setEnabled(False)
        self.confirm_btn.setEnabled(False)
        self.confirmed.emit(key)


class SuggestionCard(QFrame):
    """EC6 — câu AI hỏi lại + chip gợi ý trả lời nhanh."""

    chip_clicked = Signal(str)

    @log_call
    def __init__(self, question: str, chips: list[str], parent=None):
        super().__init__(parent)
        self.setObjectName("suggestionCard")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 8)
        lay.setSpacing(8)

        bubble_row = QHBoxLayout()
        bubble_row.setContentsMargins(0, 0, 0, 0)
        bubble = QLabel(question)
        bubble.setProperty("class", "bubbleAI")
        bubble.setWordWrap(True)
        bubble.setMaximumWidth(430)
        bubble_row.addWidget(bubble)
        bubble_row.addStretch()
        lay.addLayout(bubble_row)

        if chips:
            self._chip_row = QHBoxLayout()
            self._chip_row.setSpacing(7)
            for label in chips:
                chip = QPushButton(label)
                chip.setProperty("class", "chip")
                chip.setCursor(Qt.CursorShape.PointingHandCursor)
                chip.clicked.connect(lambda _, chip_label=label: self._on_chip(chip_label))
                self._chip_row.addWidget(chip)
            self._chip_row.addStretch()
            lay.addLayout(self._chip_row)
        else:
            self._chip_row = None

    @log_call
    def _on_chip(self, text: str):
        if self._chip_row is not None:
            for chip_index in range(self._chip_row.count()):
                chip_widget = self._chip_row.itemAt(chip_index).widget()
                if chip_widget:
                    chip_widget.setEnabled(False)
        self.chip_clicked.emit(text)


class QuotaWarningCard(QFrame):
    """EC4 — hết tác vụ AI trong tháng."""

    upgrade_clicked = Signal()
    dismissed = Signal()

    @log_call
    def __init__(self, message: str, hint: str = "", parent=None):
        super().__init__(parent)
        self.setObjectName("quotaWarningCard")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(14, 13, 14, 13)
        lay.setSpacing(6)

        title = QLabel(f"⚠ {message}")
        title.setObjectName("quotaWarningTitle")
        title.setWordWrap(True)
        lay.addWidget(title)

        if hint:
            sub = QLabel(hint)
            sub.setObjectName("quotaWarningHint")
            sub.setWordWrap(True)
            lay.addWidget(sub)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        upgrade = QPushButton(S.UPGRADE_PRO)
        upgrade.setProperty("class", "primaryBtn")
        upgrade.setCursor(Qt.CursorShape.PointingHandCursor)
        upgrade.clicked.connect(self._on_upgrade)
        btn_row.addWidget(upgrade)
        later = QPushButton(S.LATER)
        later.setProperty("class", "secondaryBtn")
        later.setCursor(Qt.CursorShape.PointingHandCursor)
        later.clicked.connect(self._on_dismiss)
        btn_row.addWidget(later)
        btn_row.addStretch()
        lay.addLayout(btn_row)

    @log_call
    def _on_upgrade(self):
        self.setEnabled(False)
        self.upgrade_clicked.emit()

    @log_call
    def _on_dismiss(self):
        self.setEnabled(False)
        self.dismissed.emit()


class GenErrorCard(QFrame):
    """EC5 — soạn thất bại giữa chừng."""

    retry_clicked = Signal()
    save_partial_clicked = Signal()

    @log_call
    def __init__(self, message: str, pages_done: int, total: int,
                 can_save_partial: bool, parent=None):
        super().__init__(parent)
        self.setObjectName("genErrorCard")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 8)
        lay.setSpacing(0)

        head = QFrame()
        head.setObjectName("genErrorHead")
        head_lay = QHBoxLayout(head)
        head_lay.setContentsMargins(13, 10, 13, 10)
        head_lay.setSpacing(9)
        icon = QLabel("✕")
        icon.setObjectName("genErrorIcon")
        icon.setFixedSize(22, 22)
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        head_lay.addWidget(icon)
        title = QLabel(S.GEN_FAILED_TITLE)
        title.setObjectName("genErrorTitle")
        head_lay.addWidget(title)
        head_lay.addStretch()
        if total:
            prog = QLabel(S.GEN_PROGRESS.format(done=pages_done, total=total))
            prog.setObjectName("genErrorProgress")
            head_lay.addWidget(prog)
        lay.addWidget(head)

        body = QLabel(message)
        body.setObjectName("genErrorBody")
        body.setWordWrap(True)
        lay.addWidget(body)

        btn_row = QFrame()
        br_lay = QHBoxLayout(btn_row)
        br_lay.setContentsMargins(13, 10, 13, 12)
        br_lay.setSpacing(8)
        retry = QPushButton(S.RETRY)
        retry.setProperty("class", "primaryBtn")
        retry.setCursor(Qt.CursorShape.PointingHandCursor)
        retry.clicked.connect(self._on_retry)
        br_lay.addWidget(retry)
        if can_save_partial:
            save_partial = QPushButton(S.SAVE_PARTIAL.format(done=pages_done))
            save_partial.setProperty("class", "secondaryBtn")
            save_partial.setCursor(Qt.CursorShape.PointingHandCursor)
            save_partial.clicked.connect(self._on_save_partial)
            br_lay.addWidget(save_partial)
        br_lay.addStretch()
        lay.addWidget(btn_row)

    @log_call
    def _on_retry(self):
        self.setEnabled(False)
        self.retry_clicked.emit()

    @log_call
    def _on_save_partial(self):
        self.setEnabled(False)
        self.save_partial_clicked.emit()
