"""Main Window — 1 cửa sổ duy nhất: sidebar · preview tài liệu · chat AI."""
from pathlib import Path
from typing import Optional

from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFrame, QSplitter, QStackedWidget, QFileDialog, QMessageBox,
)

from ..ai import classify_intent, suggest_save_name, MockReply
from ..config import load_config, save_config
from ..constants import (
    APP_NAME, EXT_MAP, FILE_BADGE, CONTEXT_CHIPS, FILE_DIALOG_FILTER,
)
from ..document import apply_edits, check_editable, document_outline, EditError
from .. import api_client
from .sidebar import Sidebar
from .empty_state import CentralChat
from .preview import PreviewPanel
from .chat import ChatPanel
from .modals import SettingsDialog, ConvertDialog, OverwriteDialog
from .toast import CacheToast
from .workers import CallWorker
from .folder_workspace import FolderWorkspace
from .folder_scan import FolderScanWorker

_MAX_RECENT = 10


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(1360, 860)
        self.setMinimumSize(1000, 640)
        self.setAcceptDrops(True)

        self._current_path: Optional[str] = None
        self._current_type: Optional[str] = None
        self._stream_timer: Optional[QTimer] = None
        self._pending_reply: Optional[MockReply] = None
        self._chat_worker: Optional[CallWorker] = None
        self._edit_worker: Optional[CallWorker] = None
        self._me_worker: Optional[CallWorker] = None
        self._pending_edit_notes: list[str] = []
        self._history: list[dict] = []   # {role, content} — ngữ cảnh cho AI
        self._plan = "free"
        self._quota_remaining: Optional[int] = None
        self._quota_limit: Optional[int] = None

        self._context_mode = "file"      # "file" | "folder"
        self._folder_path: Optional[str] = None
        self._folder_files: list = []
        self._folder_selected: list[str] = []
        self._folder_context_name: Optional[str] = None
        self._folder_scan_worker: Optional[FolderScanWorker] = None

        self._build_ui()
        self._cache_toast = CacheToast(self)
        self._refresh_sidebar()
        self._update_quota_label()
        self._refresh_account()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self._cache_toast.isVisible():
            self._cache_toast.reposition()

    # ══ UI ═════════════════════════════════════════════════════════════════

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        root.addWidget(self._make_topbar())

        body = QHBoxLayout()
        body.setContentsMargins(0, 0, 0, 0)
        body.setSpacing(0)

        self.sidebar = Sidebar()
        self.sidebar.new_chat.connect(self._new_chat)
        self.sidebar.file_selected.connect(self._open_file)
        self.sidebar.mode_changed.connect(self._on_mode_changed)
        self.sidebar.open_folder.connect(self._open_folder)
        self.sidebar.change_folder.connect(self._open_folder)
        self.sidebar.folder_selection_changed.connect(self._on_folder_selection)
        self.sidebar.recent_deleted.connect(self._on_recent_deleted)
        self.sidebar.folder_file_removed.connect(self._on_folder_file_removed)
        body.addWidget(self.sidebar)

        # Stack: [0] chat AI trung tâm (mở app là sẵn sàng) · [1] workspace (preview + chat)
        self.stack = QStackedWidget()

        self.welcome = CentralChat()
        self.welcome.file_chosen.connect(self._open_file)
        self.welcome.message_sent.connect(self._start_chat)
        self.stack.addWidget(self.welcome)

        workspace = QWidget()
        ws_lay = QHBoxLayout(workspace)
        ws_lay.setContentsMargins(10, 10, 10, 10)
        ws_lay.setSpacing(0)

        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.splitter.setChildrenCollapsible(False)

        # Trạng thái 2 (theo MVP §4.1): cột chat thu hẹp bên TRÁI,
        # panel xem trước/chỉnh sửa file bên PHẢI.
        self.chat = ChatPanel()
        self.chat.message_sent.connect(self._on_message)
        self.chat.file_card_clicked.connect(self._on_file_card)
        self.chat.attach_requested.connect(self._pick_file)
        self.splitter.addWidget(self.chat)

        self.preview = PreviewPanel()
        self.preview.convert_requested.connect(self._open_convert)
        self.splitter.addWidget(self.preview)

        self.splitter.setStretchFactor(0, 0)   # chat: giữ hẹp
        self.splitter.setStretchFactor(1, 1)   # preview: co giãn
        self.splitter.setSizes([360, 720])

        ws_lay.addWidget(self.splitter)
        self.stack.addWidget(workspace)

        self.folder_view = FolderWorkspace()
        self.folder_view.open_requested.connect(self._open_folder)
        self.folder_view.select_all_requested.connect(self._folder_select_all)
        self.folder_view.start_chat_requested.connect(self._start_folder_chat)
        self.stack.addWidget(self.folder_view)

        body.addWidget(self.stack, stretch=1)
        root.addLayout(body, stretch=1)

    def _make_topbar(self) -> QFrame:
        bar = QFrame()
        bar.setObjectName("topbar")
        bar.setFixedHeight(52)
        lay = QHBoxLayout(bar)
        lay.setContentsMargins(16, 0, 12, 0)
        lay.setSpacing(9)

        logo = QLabel("D")
        logo.setObjectName("appLogoTile")
        logo.setFixedSize(26, 26)
        logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(logo)

        title = QLabel(APP_NAME)
        title.setObjectName("appTitle")
        lay.addWidget(title)
        lay.addStretch()

        self.quota_lbl = QLabel("")
        self.quota_lbl.setObjectName("quotaLabel")
        lay.addWidget(self.quota_lbl)

        gear = QPushButton("⚙")
        gear.setObjectName("gearBtn")
        gear.setCursor(Qt.CursorShape.PointingHandCursor)
        gear.clicked.connect(self._open_settings)
        lay.addWidget(gear)
        return bar

    # ══ Kéo-thả file vào cửa sổ ══════════════════════════════════════════════

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            path = event.mimeData().urls()[0].toLocalFile()
            if Path(path).suffix.lower() in EXT_MAP:
                event.acceptProposedAction()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if Path(path).suffix.lower() in EXT_MAP:
                self._open_file(path)
                return

    # ══ Mở file / luồng bắt đầu ══════════════════════════════════════════════

    def _pick_file(self):
        """Đính kèm file từ ô chat workspace — gắn vào cuộc trò chuyện đang có."""
        path, _ = QFileDialog.getOpenFileName(self, "Chọn file", "", FILE_DIALOG_FILTER)
        if path:
            self._open_file(path, fresh=False)

    def _open_file(self, path: str, fresh: bool = True):
        p = Path(path)
        if not p.exists():
            QMessageBox.warning(self, APP_NAME, f"File không tồn tại:\n{path}")
            self._remove_recent(path)
            return
        ftype = EXT_MAP.get(p.suffix.lower())
        if ftype is None:
            QMessageBox.warning(self, APP_NAME, f"Định dạng chưa hỗ trợ: {p.suffix}")
            return

        self._current_path = str(p)
        self._current_type = ftype
        self._push_recent(str(p))
        self._refresh_sidebar()

        if fresh:
            self.chat.clear()
            self._history.clear()
            self._folder_context_name = None

        # Trạng thái 2 — chat thu hẹp bên trái, panel file bên phải.
        self.preview.setVisible(True)
        self.preview.load_file(str(p), ftype, p.name)
        self.chat.set_chips(CONTEXT_CHIPS.get(ftype, CONTEXT_CHIPS[None]))
        self.stack.setCurrentIndex(1)
        verb = "Đã mở" if fresh else "Đã đính kèm"
        self.chat.add_system(f"{verb} «{p.name}»")
        self.chat.input_box.setFocus()

    def _start_chat(self, text: str):
        """Gõ lệnh / chip ở chat trung tâm → vào workspace, chưa cần file."""
        self._current_path = None
        self._current_type = None
        self.chat.clear()
        self.preview.setVisible(False)
        self.chat.set_chips(CONTEXT_CHIPS[None])
        self.stack.setCurrentIndex(1)
        self.chat.send_text(text)

    def _on_mode_changed(self, mode: str):
        """Chuyển chế độ nạp ngữ cảnh Tệp ↔ Thư mục."""
        self._context_mode = mode
        if mode == "folder":
            self.stack.setCurrentIndex(2)
        else:
            self.stack.setCurrentIndex(1 if self._current_path else 0)

    def _new_chat(self):
        """Trò chuyện mới → về chat AI trung tâm, xóa hội thoại & ngữ cảnh file."""
        self._current_path = None
        self._current_type = None
        self._folder_context_name = None
        self.chat.clear()
        self._history.clear()
        self.preview.setVisible(False)
        self._refresh_sidebar()
        self.stack.setCurrentIndex(0)
        self.welcome.input_box.setFocus()

    # ══ Chế độ Thư mục — mở / quét / chọn file làm ngữ cảnh ══════════════════

    def _open_folder(self):
        path = QFileDialog.getExistingDirectory(self, "Chọn thư mục làm việc", "")
        if not path:
            return

        self._folder_path = path
        self._folder_files = []
        self._folder_selected = []
        folder_name = Path(path).name or path

        self.sidebar.show_folder_scanning(path)
        self.folder_view.show_scanning(folder_name)
        self.stack.setCurrentIndex(2)

        self._folder_scan_worker = FolderScanWorker(path)
        self._folder_scan_worker.progress.connect(self._on_folder_scan_progress)
        self._folder_scan_worker.done.connect(self._on_folder_scan_done)
        self._folder_scan_worker.start()

    def _on_folder_scan_progress(self, done: int, total: int, counts: dict):
        self.sidebar.update_folder_scan_progress(done, total)
        self.folder_view.update_progress(done, total, counts)

    def _on_folder_scan_done(self, files: list, counts: dict):
        self._folder_files = files
        folder_name = Path(self._folder_path).name or self._folder_path
        self.sidebar.show_folder_ready(folder_name, files, counts)
        self.folder_view.show_ready(folder_name, len(files))

    def _on_folder_selection(self, paths: list[str]):
        self._folder_selected = paths

    def _on_folder_file_removed(self, path: str):
        """Bấm xóa 1 file trong cây thư mục — bỏ khỏi danh sách/ngữ cảnh, không đụng file thật."""
        self._folder_files = [f for f in self._folder_files if f.path != path]
        self._folder_selected = [p for p in self._folder_selected if p != path]

    def _folder_select_all(self):
        self.sidebar.select_all_folder_files()

    def _start_folder_chat(self):
        selected = self._folder_selected or [f.path for f in self._folder_files]
        if not selected or not self._folder_path:
            return

        self._current_path = None
        self._current_type = None
        self.chat.clear()
        self._history.clear()
        self.preview.setVisible(False)
        self.chat.set_chips(CONTEXT_CHIPS[None])

        folder_name = Path(self._folder_path).name or self._folder_path
        self._folder_context_name = f"{folder_name} ({len(selected)} file)"
        self.stack.setCurrentIndex(1)
        self.chat.add_system(
            f"Đã mở thư mục «{folder_name}» — {len(selected)} file được chọn làm ngữ cảnh")
        self.chat.input_box.setFocus()

    # ══ Sidebar / recent ═════════════════════════════════════════════════════

    def _push_recent(self, path: str):
        cfg = load_config()
        recent = cfg.get("recent_files", [])
        if path not in recent:
            recent = [path] + recent
        cfg["recent_files"] = recent[:_MAX_RECENT]
        save_config(cfg)

    def _remove_recent(self, path: str):
        cfg = load_config()
        cfg["recent_files"] = [r for r in cfg.get("recent_files", []) if r != path]
        save_config(cfg)
        self._refresh_sidebar()

    def _on_recent_deleted(self, path: str):
        """Bấm xóa 1 item trong tab "Tệp" — chỉ gỡ khỏi lịch sử, không xóa file thật."""
        self._remove_recent(path)

    def _refresh_sidebar(self):
        recent = load_config().get("recent_files", [])
        self.sidebar.refresh(recent, self._current_path)

    # ══ Quota (nguồn: server) ════════════════════════════════════════════════

    def _update_quota_label(self):
        plan = "Pro" if self._plan == "pro" else "Free"
        if self._quota_remaining is None or self._quota_limit is None:
            self.quota_lbl.setText(f"Gói {plan}")
        else:
            self.quota_lbl.setText(
                f"Gói {plan} · {self._quota_remaining}/{self._quota_limit} lượt")
        self.sidebar.set_plan(plan, self._quota_remaining, self._quota_limit)

    def _refresh_account(self):
        """Lấy gói + quota từ server (nền)."""
        self._me_worker = CallWorker(api_client.me)
        self._me_worker.ok.connect(self._on_account)
        self._me_worker.err.connect(lambda _msg: None)  # im lặng nếu lỗi mạng
        self._me_worker.start()

    def _on_account(self, data: dict):
        self._plan = data.get("plan", "free")
        self._quota_remaining = data.get("quota_remaining")
        self._quota_limit = data.get("quota_limit")
        self._update_quota_label()

    # ══ Chat AI (gọi server, streaming khi có kết quả) ═══════════════════════

    def _on_message(self, text: str):
        self.chat.add_user(text)
        self.chat.set_enabled(False)
        self._remember("user", text)

        # Đang mở file Word/Excel → câu chat là lệnh sửa file (nếu AI hiểu vậy).
        if self._current_path and self._current_type in ("word", "excel"):
            self._start_edit(text)
            return

        file_name = (
            Path(self._current_path).name if self._current_path
            else self._folder_context_name)
        business = load_config().get("business", {})

        # Meta (file-card / ghi đè) suy từ câu lệnh; text lấy từ server.
        artifact, is_edit = classify_intent(text, file_name, self._current_type)
        self._pending_meta = (artifact, is_edit)

        self.chat.start_ai()
        self.chat.stream_ai("Đang xử lý…")

        self._chat_worker = CallWorker(
            api_client.chat, text, file_name or "", self._current_type or "",
            business,
        )
        self._chat_worker.ok.connect(self._on_chat_result)
        self._chat_worker.err.connect(self._on_chat_error)
        self._chat_worker.start()

    def _on_chat_result(self, data: dict):
        # Cập nhật quota từ server.
        if data.get("quota_remaining") is not None:
            self._quota_remaining = data["quota_remaining"]
            self._plan = data.get("plan", self._plan)
            self._update_quota_label()
        self._cache_toast.show_usage(data.get("usage"))

        artifact, is_edit = getattr(self, "_pending_meta", (None, False))
        self._pending_reply = MockReply(
            text=data.get("text", "") or "(Không có nội dung trả về)",
            artifact=artifact, is_edit=is_edit,
        )
        self._start_stream(self._pending_reply.text)

    def _start_stream(self, full_text: str):
        """Hiển thị dần văn bản server trả về."""
        self._stream_pos = 0
        self._stream_timer = QTimer(self)
        self._stream_timer.setInterval(16)
        self._stream_timer.timeout.connect(lambda: self._stream_tick(full_text))
        self._stream_timer.start()

    def _on_chat_error(self, message: str):
        self.chat.stream_ai("")
        self.chat.end_ai()
        self.chat.add_system(f"⚠ {message}")
        self.chat.set_enabled(True)
        self.chat.input_box.setFocus()
        self._pending_reply = None

    def _stream_tick(self, full_text: str):
        self._stream_pos = min(len(full_text), self._stream_pos + 6)
        self.chat.stream_ai(full_text[:self._stream_pos])
        if self._stream_pos >= len(full_text):
            self._stream_timer.stop()
            self._finish_reply()

    # ══ Sửa file thật (Word / Excel) ═════════════════════════════════════════

    def _remember(self, role: str, content: str):
        self._history.append({"role": role, "content": content})
        del self._history[:-12]

    def _start_edit(self, text: str):
        path, ftype = self._current_path, self._current_type
        try:
            check_editable(path, ftype)
        except EditError as exc:
            self.chat.add_system(f"⚠ {exc}")
            self.chat.set_enabled(True)
            self.chat.input_box.setFocus()
            return

        self.chat.start_ai()
        self.chat.stream_ai("Đang xử lý…")

        self._edit_worker = CallWorker(
            self._request_edit, text, path, ftype, list(self._history[:-1]))
        self._edit_worker.ok.connect(self._on_edit_result)
        self._edit_worker.err.connect(self._on_chat_error)
        self._edit_worker.start()

    @staticmethod
    def _request_edit(text: str, path: str, ftype: str, history: list) -> dict:
        """Chạy ở luồng nền: gửi cấu trúc tài liệu + yêu cầu lên server, rồi áp
        các lệnh sửa Claude trả về. `edits` rỗng → không phải yêu cầu sửa."""
        try:
            outline = document_outline(path, ftype)
        except EditError as exc:
            raise api_client.ApiError(str(exc), "outline_failed")

        data = api_client.edit_file(text, Path(path).name, ftype, outline, history)

        edits = data.get("edits") or []
        if not edits:
            return {**data, "notes": []}
        try:
            notes = apply_edits(path, ftype, edits)
        except EditError as exc:
            raise api_client.ApiError(str(exc), "edit_failed")
        return {**data, "notes": notes}

    def _on_edit_result(self, data: dict):
        if data.get("quota_remaining") is not None:
            self._quota_remaining = data["quota_remaining"]
            self._plan = data.get("plan", self._plan)
            self._update_quota_label()
        self._cache_toast.show_usage(data.get("usage"))

        self._pending_meta = (None, False)
        self._pending_edit_notes = data.get("notes") or []
        self._pending_reply = MockReply(
            text=data.get("reply", "") or "Đã cập nhật file.")
        self._start_stream(self._pending_reply.text)

    def _finish_reply(self):
        self.chat.end_ai()
        reply = self._pending_reply
        self._pending_reply = None
        self.chat.set_enabled(True)
        self.chat.input_box.setFocus()

        if reply is not None and reply.text:
            self._remember("assistant", reply.text)

        # File vừa được sửa thật → báo từng thay đổi & render lại bản xem trước.
        notes, self._pending_edit_notes = self._pending_edit_notes, []
        if notes and self._current_path:
            name = Path(self._current_path).name
            for note in notes:
                self.chat.add_system(f"✓ «{name}» — {note}")
            self.preview.load_file(self._current_path, self._current_type, name)
            return

        if reply is None:
            return

        if reply.artifact:
            badge = FILE_BADGE.get(EXT_MAP.get(Path(reply.artifact).suffix.lower(), ""), "FILE")
            self.chat.add_file_card(reply.artifact, badge)

        if reply.is_edit and self._current_path:
            self._confirm_overwrite()

    # ══ Modal: Xác nhận ghi đè ═══════════════════════════════════════════════

    def _confirm_overwrite(self):
        name = Path(self._current_path).name
        dlg = OverwriteDialog(name, self)
        if dlg.exec():
            if dlg.choice == OverwriteDialog.OVERWRITE:
                self.chat.add_system(f"✓ Đã ghi đè thay đổi lên «{name}»")
            else:
                copy_name = suggest_save_name(name)
                badge = FILE_BADGE.get(self._current_type or "", "FILE")
                self.chat.add_system("✓ Đã lưu bản sao, giữ nguyên file gốc")
                self.chat.add_file_card(copy_name, badge)

    # ══ Modal: Chuyển đổi định dạng ══════════════════════════════════════════

    def _open_convert(self):
        if not self._current_path:
            return
        name = Path(self._current_path).name
        dlg = ConvertDialog(name, self._current_type or "pdf", self)
        dlg.conversion_done.connect(self._on_converted)
        dlg.exec()

    def _on_converted(self, out_name: str):
        badge = FILE_BADGE.get(EXT_MAP.get(Path(out_name).suffix.lower(), ""), "FILE")
        self.chat.add_system("✓ Chuyển đổi định dạng hoàn tất")
        self.chat.add_file_card(out_name, badge)

    # ══ Modal: Cài đặt ═══════════════════════════════════════════════════════

    def _open_settings(self):
        dlg = SettingsDialog(self)
        dlg.logged_out.connect(self._handle_logout)
        dlg.exec()
        self._refresh_account()

    def _handle_logout(self):
        from ..app import restart_to_login
        restart_to_login(self)

    # ══ File card (mock) ═════════════════════════════════════════════════════

    def _on_file_card(self, file_name: str):
        QMessageBox.information(
            self, APP_NAME,
            f"«{file_name}»\n\nFile kết quả sẽ được tạo thật khi tích hợp backend AI.")

    # ══ Dọn dẹp ══════════════════════════════════════════════════════════════

    def closeEvent(self, event):
        self.preview.cleanup()
        super().closeEvent(event)
