"""Splash — Đăng nhập / Đăng ký tài khoản DocAI (Firebase Auth)."""
import re

from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtWidgets import (
    QWidget, QFrame, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
)

from ..constants import APP_NAME, APP_TAGLINE
from .. import api_client
from ..server_config import is_configured
from .workers import CallWorker

_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def detect_ms_word() -> bool:
    """Kiểm tra máy có đăng ký Word COM không (đọc registry, không mở Word)."""
    try:
        import winreg
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Word.Application"))
        return True
    except OSError:
        return False


class SplashWindow(QWidget):
    def __init__(self, on_activated):
        super().__init__()
        self._on_activated = on_activated
        self._mode = "login"          # "login" | "signup"
        self._worker: CallWorker | None = None
        self.setWindowTitle(f"{APP_NAME} — Đăng nhập")
        self.setFixedSize(620, 620)
        self._build_ui()
        QTimer.singleShot(400, self._check_word)

    # ── UI ─────────────────────────────────────────────────────────────────

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(80, 50, 80, 50)

        card = QFrame()
        card.setObjectName("splashCard")
        lay = QVBoxLayout(card)
        lay.setContentsMargins(48, 40, 48, 32)
        lay.setSpacing(10)

        logo = QLabel("D")
        logo.setObjectName("splashLogo")
        logo.setFixedSize(56, 56)
        logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(logo, alignment=Qt.AlignmentFlag.AlignHCenter)

        title = QLabel(APP_NAME)
        title.setObjectName("splashTitle")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(title)

        sub = QLabel(APP_TAGLINE)
        sub.setObjectName("splashSub")
        sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(sub)
        lay.addSpacing(18)

        email_lbl = QLabel("Email")
        email_lbl.setObjectName("fieldLabel")
        lay.addWidget(email_lbl)
        self.email_input = QLineEdit()
        self.email_input.setPlaceholderText("ban@congty.vn")
        lay.addWidget(self.email_input)

        pw_lbl = QLabel("Mật khẩu")
        pw_lbl.setObjectName("fieldLabel")
        lay.addWidget(pw_lbl)
        self.pw_input = QLineEdit()
        self.pw_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.pw_input.setPlaceholderText("Tối thiểu 6 ký tự")
        self.pw_input.returnPressed.connect(self._submit)
        lay.addWidget(self.pw_input)

        self.error_lbl = QLabel("")
        self.error_lbl.setObjectName("splashError")
        self.error_lbl.setWordWrap(True)
        self.error_lbl.setVisible(False)
        lay.addWidget(self.error_lbl)
        lay.addSpacing(4)

        self.submit_btn = QPushButton("Đăng nhập")
        self.submit_btn.setProperty("class", "primaryBtn")
        self.submit_btn.setFixedHeight(42)
        self.submit_btn.clicked.connect(self._submit)
        lay.addWidget(self.submit_btn)

        # Chuyển đổi Đăng nhập ↔ Đăng ký
        switch_row = QHBoxLayout()
        switch_row.addStretch()
        self.switch_hint = QLabel("Chưa có tài khoản?")
        self.switch_hint.setObjectName("splashSub")
        switch_row.addWidget(self.switch_hint)
        self.switch_btn = QPushButton("Đăng ký")
        self.switch_btn.setProperty("class", "linkBtn")
        self.switch_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.switch_btn.clicked.connect(self._toggle_mode)
        switch_row.addWidget(self.switch_btn)
        switch_row.addStretch()
        lay.addLayout(switch_row)

        lay.addSpacing(6)
        self.status_lbl = QLabel("Đang kiểm tra Microsoft Word…")
        self.status_lbl.setObjectName("splashStatus")
        self.status_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_lbl.setWordWrap(True)
        lay.addWidget(self.status_lbl)

        root.addWidget(card)

        if not is_configured():
            self._show_error(
                "Chưa cấu hình server: điền FIREBASE_API_KEY và API_BASE_URL "
                "trong server_config.py (hoặc config.json).")

    def _toggle_mode(self):
        self._mode = "signup" if self._mode == "login" else "login"
        if self._mode == "signup":
            self.submit_btn.setText("Đăng ký")
            self.switch_hint.setText("Đã có tài khoản?")
            self.switch_btn.setText("Đăng nhập")
        else:
            self.submit_btn.setText("Đăng nhập")
            self.switch_hint.setText("Chưa có tài khoản?")
            self.switch_btn.setText("Đăng ký")
        self.error_lbl.setVisible(False)

    # ── Kiểm tra Word COM ────────────────────────────────────────────────────

    def _check_word(self):
        has_word = detect_ms_word()
        from ..config import load_config, save_config
        cfg = load_config()
        cfg["office_engine"] = "word_com" if has_word else "libreoffice"
        save_config(cfg)
        self.status_lbl.setText(
            "Đã tìm thấy Microsoft Word — dùng Word COM để xử lý tài liệu."
            if has_word else
            "Không thấy Microsoft Word — sẽ dùng LibreOffice dự phòng."
        )

    # ── Đăng nhập / Đăng ký ──────────────────────────────────────────────────

    def _show_error(self, msg: str):
        self.error_lbl.setText(msg)
        self.error_lbl.setVisible(True)

    def _set_busy(self, busy: bool):
        self.submit_btn.setEnabled(not busy)
        self.switch_btn.setEnabled(not busy)
        self.email_input.setEnabled(not busy)
        self.pw_input.setEnabled(not busy)
        if busy:
            self.submit_btn.setText(
                "Đang đăng nhập…" if self._mode == "login" else "Đang đăng ký…")
        else:
            self.submit_btn.setText(
                "Đăng nhập" if self._mode == "login" else "Đăng ký")

    def _submit(self):
        email = self.email_input.text().strip()
        password = self.pw_input.text()
        if not _EMAIL_RE.match(email):
            self._show_error("Email không hợp lệ.")
            return
        if len(password) < 6:
            self._show_error("Mật khẩu cần tối thiểu 6 ký tự.")
            return

        self.error_lbl.setVisible(False)
        self._set_busy(True)
        fn = api_client.login if self._mode == "login" else api_client.signup
        self._worker = CallWorker(fn, email, password)
        self._worker.ok.connect(self._on_success)
        self._worker.err.connect(self._on_error)
        self._worker.start()

    def _on_success(self, _result):
        self._set_busy(False)
        self._on_activated()
        self.close()

    def _on_error(self, message: str):
        self._set_busy(False)
        self._show_error(message)
