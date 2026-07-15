"""3 modal phụ: Cài đặt & tài khoản · Chuyển đổi định dạng · Xác nhận ghi đè."""
from pathlib import Path

from PyQt6.QtCore import Qt, QTimer, pyqtSignal
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QLineEdit,
    QComboBox, QProgressBar, QFrame, QStackedWidget, QWidget, QFormLayout,
)

from ..config import load_config, save_config
from ..constants import CONVERT_TARGETS, CONVERT_EXT
from .. import api_client
from .workers import CallWorker


def _modal_header(dialog: QDialog, title: str) -> QHBoxLayout:
    row = QHBoxLayout()
    lbl = QLabel(title)
    lbl.setObjectName("modalTitle")
    row.addWidget(lbl)
    row.addStretch()
    close = QPushButton("✕")
    close.setObjectName("modalClose")
    close.setCursor(Qt.CursorShape.PointingHandCursor)
    close.clicked.connect(dialog.reject)
    row.addWidget(close)
    return row


# ══ 4.4 Cài đặt & tài khoản ══════════════════════════════════════════════════

class SettingsDialog(QDialog):
    logged_out = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Cài đặt & tài khoản")
        self.setModal(True)
        self.setFixedSize(640, 560)
        self.setStyleSheet("QDialog { background: white; }")

        self._me_worker: CallWorker | None = None
        self._biz_worker: CallWorker | None = None
        cfg = load_config()
        lay = QVBoxLayout(self)
        lay.setContentsMargins(28, 24, 28, 20)
        lay.setSpacing(14)
        lay.addLayout(_modal_header(self, "Cài đặt & tài khoản"))

        # ── Tab buttons ───────────────────────────────────────────────────────
        tabs_row = QHBoxLayout()
        tabs_row.setSpacing(8)
        self._tab_btns: list[QPushButton] = []
        for i, name in enumerate(["Tài khoản", "Doanh nghiệp", "Giới hạn"]):
            btn = QPushButton(name)
            btn.setProperty("class", "tabBtn")
            btn.setCheckable(True)
            btn.setChecked(i == 0)
            btn.clicked.connect(lambda _, idx=i: self._switch_tab(idx))
            self._tab_btns.append(btn)
            tabs_row.addWidget(btn)
        tabs_row.addStretch()
        lay.addLayout(tabs_row)

        self._stack = QStackedWidget()
        self._stack.addWidget(self._build_account_tab(cfg))
        self._stack.addWidget(self._build_business_tab(cfg))
        self._stack.addWidget(self._build_limits_tab(cfg))
        lay.addWidget(self._stack, stretch=1)

        # ── Footer ────────────────────────────────────────────────────────────
        save_row = QHBoxLayout()
        save_row.addStretch()
        save_btn = QPushButton("Lưu thay đổi")
        save_btn.setProperty("class", "primaryBtn")
        save_btn.clicked.connect(self._save)
        save_row.addWidget(save_btn)
        lay.addLayout(save_row)

        note = QLabel("* Thông tin doanh nghiệp dùng để AI tự điền khi soạn văn bản hành chính")
        note.setObjectName("footnote")
        lay.addWidget(note)

        self._load_account()

    def _switch_tab(self, idx: int):
        for i, btn in enumerate(self._tab_btns):
            btn.setChecked(i == idx)
        self._stack.setCurrentIndex(idx)

    # ── Tab Tài khoản ─────────────────────────────────────────────────────────

    def _build_account_tab(self, cfg: dict) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(0, 10, 0, 0)
        lay.setSpacing(12)

        self._acc_vals: dict[str, QLabel] = {}
        for key, label, initial in [
            ("email", "Tài khoản", api_client.current_email() or "—"),
            ("plan", "Gói hiện tại", "Đang tải…"),
            ("quota", "Quota còn lại", "Đang tải…"),
        ]:
            row = QHBoxLayout()
            lbl = QLabel(label)
            lbl.setObjectName("fieldLabel")
            row.addWidget(lbl)
            row.addStretch()
            val = QLabel(initial)
            val.setObjectName("valueLabel")
            self._acc_vals[key] = val
            row.addWidget(val)
            lay.addLayout(row)

        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setStyleSheet("color: #DDDCD4;")
        lay.addWidget(line)
        lay.addStretch()

        logout_row = QHBoxLayout()
        logout_row.addStretch()
        logout_btn = QPushButton("Đăng xuất")
        logout_btn.setProperty("class", "secondaryBtn")
        logout_btn.clicked.connect(self._logout)
        logout_row.addWidget(logout_btn)
        lay.addLayout(logout_row)
        return w

    def _load_account(self):
        self._me_worker = CallWorker(api_client.me)
        self._me_worker.ok.connect(self._on_account)
        self._me_worker.err.connect(lambda msg: self._acc_vals["plan"].setText(msg))
        self._me_worker.start()

    def _on_account(self, data: dict):
        plan = "Pro" if data.get("plan") == "pro" else "Free"
        self._acc_vals["email"].setText(data.get("email") or "—")
        self._acc_vals["plan"].setText(plan)
        self._acc_vals["quota"].setText(
            f"{data.get('quota_remaining', 0)} / {data.get('quota_limit', 0)} lượt")
        # Đồng bộ thông tin doanh nghiệp từ server vào ô nhập nếu có.
        biz = data.get("business") or {}
        for k, edit in getattr(self, "biz_inputs", {}).items():
            if biz.get(k) and not edit.text().strip():
                edit.setText(str(biz[k]))

    def _logout(self):
        api_client.logout()
        self.logged_out.emit()
        self.accept()

    # ── Tab Doanh nghiệp ──────────────────────────────────────────────────────

    def _build_business_tab(self, cfg: dict) -> QWidget:
        w = QWidget()
        form = QFormLayout(w)
        form.setContentsMargins(0, 10, 0, 0)
        form.setVerticalSpacing(12)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignLeft)

        biz = cfg.get("business", {})
        self.biz_inputs: dict[str, QLineEdit] = {}
        for key, label in [
            ("company_name", "Tên doanh nghiệp"),
            ("tax_code", "Mã số thuế (MST)"),
            ("address", "Địa chỉ"),
            ("representative", "Người đại diện"),
        ]:
            lbl = QLabel(label)
            lbl.setObjectName("fieldLabel")
            edit = QLineEdit(biz.get(key, ""))
            self.biz_inputs[key] = edit
            form.addRow(lbl, edit)
        return w

    # ── Tab Giới hạn ──────────────────────────────────────────────────────────

    def _build_limits_tab(self, cfg: dict) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(0, 10, 0, 0)
        lbl = QLabel("Chi tiết quota theo gói sẽ có ở Phase 2.")
        lbl.setObjectName("fieldLabel")
        lay.addWidget(lbl)
        lay.addStretch()
        return w

    def _save(self):
        business = {k: e.text().strip() for k, e in self.biz_inputs.items()}
        cfg = load_config()
        cfg["business"] = business   # lưu cục bộ để chat dùng ngay
        save_config(cfg)
        # Đồng bộ lên server (nền); không chặn đóng dialog.
        self._biz_worker = CallWorker(api_client.update_business, business)
        self._biz_worker.err.connect(lambda _msg: None)
        self._biz_worker.start()
        self.accept()


# ══ 4.5 Chuyển đổi định dạng ═════════════════════════════════════════════════

class ConvertDialog(QDialog):
    conversion_done = pyqtSignal(str)   # tên file kết quả

    def __init__(self, file_name: str, file_type: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Chuyển đổi định dạng")
        self.setModal(True)
        self.setFixedSize(500, 300)
        self.setStyleSheet("QDialog { background: white; }")
        self._file_name = file_name
        self._timer: QTimer | None = None

        lay = QVBoxLayout(self)
        lay.setContentsMargins(28, 24, 28, 20)
        lay.setSpacing(12)
        lay.addLayout(_modal_header(self, "Chuyển đổi định dạng"))

        src = QLabel(file_name)
        src.setObjectName("fieldLabel")
        lay.addWidget(src)
        lay.addSpacing(6)

        target_lbl = QLabel("Định dạng đích")
        target_lbl.setObjectName("fieldLabel")
        lay.addWidget(target_lbl)

        self.target_combo = QComboBox()
        self.target_combo.addItems(CONVERT_TARGETS.get(file_type, ["PDF (.pdf)"]))
        lay.addWidget(self.target_combo)

        self.status_lbl = QLabel("")
        self.status_lbl.setObjectName("fieldLabel")
        self.status_lbl.setVisible(False)
        lay.addWidget(self.status_lbl)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setFixedHeight(8)
        self.progress.setVisible(False)
        lay.addWidget(self.progress)
        lay.addStretch()

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.cancel_btn = QPushButton("Hủy")
        self.cancel_btn.setProperty("class", "secondaryBtn")
        self.cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(self.cancel_btn)
        self.convert_btn = QPushButton("Chuyển đổi")
        self.convert_btn.setProperty("class", "primaryBtn")
        self.convert_btn.clicked.connect(self._start)
        btn_row.addWidget(self.convert_btn)
        lay.addLayout(btn_row)

    def _start(self):
        # Mock: giả lập tiến trình chuyển đổi (Word COM/LibreOffice ở bước sau)
        self.convert_btn.setEnabled(False)
        self.target_combo.setEnabled(False)
        self.progress.setVisible(True)
        self.status_lbl.setVisible(True)
        self._pct = 0

        self._timer = QTimer(self)
        self._timer.timeout.connect(self._tick)
        self._timer.start(35)

    def _tick(self):
        self._pct = min(100, self._pct + 2)
        self.progress.setValue(self._pct)
        self.status_lbl.setText(f"Đang chuyển đổi… {self._pct}%")
        if self._pct >= 100:
            self._timer.stop()
            ext = CONVERT_EXT[self.target_combo.currentText()]
            out_name = Path(self._file_name).stem + ext
            self.conversion_done.emit(out_name)
            self.accept()

    def reject(self):
        if self._timer:
            self._timer.stop()
        super().reject()


# ══ 4.6 Xác nhận ghi đè file ═════════════════════════════════════════════════

class OverwriteDialog(QDialog):
    SAVE_COPY = 1
    OVERWRITE = 2

    def __init__(self, file_name: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Xác nhận ghi đè")
        self.setModal(True)
        self.setFixedSize(480, 230)
        self.setStyleSheet("QDialog { background: white; }")
        self.choice = self.SAVE_COPY

        lay = QVBoxLayout(self)
        lay.setContentsMargins(28, 26, 28, 22)
        lay.setSpacing(8)

        head = QHBoxLayout()
        head.setSpacing(12)
        icon = QLabel("!")
        icon.setObjectName("dangerIcon")
        icon.setFixedSize(40, 28)
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        head.addWidget(icon)
        title = QLabel("File sẽ bị ghi đè")
        title.setObjectName("modalTitle")
        head.addWidget(title)
        head.addStretch()
        lay.addLayout(head)

        msg = QLabel(
            f"Bạn có muốn lưu một bản sao trước khi\nAI ghi đè lên «{file_name}» không?")
        msg.setObjectName("fieldLabel")
        msg.setWordWrap(True)
        lay.addWidget(msg)
        lay.addStretch()

        btn_row = QHBoxLayout()
        btn_row.setSpacing(12)
        save_copy = QPushButton("Lưu bản sao")
        save_copy.setProperty("class", "secondaryBtn")
        save_copy.clicked.connect(self._pick_copy)
        btn_row.addWidget(save_copy, stretch=1)
        overwrite = QPushButton("Ghi đè")
        overwrite.setProperty("class", "primaryBtn")
        overwrite.clicked.connect(self._pick_overwrite)
        btn_row.addWidget(overwrite, stretch=1)
        lay.addLayout(btn_row)

    def _pick_copy(self):
        self.choice = self.SAVE_COPY
        self.accept()

    def _pick_overwrite(self):
        self.choice = self.OVERWRITE
        self.accept()
