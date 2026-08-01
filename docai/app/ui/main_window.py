"""Main Window — 1 cửa sổ duy nhất: sidebar · preview tài liệu · chat AI."""
import os
import subprocess
import sys
import shutil
import tempfile
from pathlib import Path
from typing import Optional

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFrame, QSplitter, QStackedWidget, QFileDialog, QMessageBox,
)

from ...account import api_client
from ...account.config import load_config, save_config
from ...modules.ai.chat import classify_intent, suggest_save_name, MockReply
from ...modules.ai.generation import (
    detect_create_intent, build_creation_prompt, suggest_file_name, guess_page_count,
)
from ...modules.common.creators import create_word, create_pptx
from ...modules.common.extractors import extract_pdf
from ...pipeline.editing import check_editable, EditError
from .. import controller
from ..constants import (
    APP_NAME, EXT_MAP, FILE_BADGE, CONTEXT_CHIPS, FILE_DIALOG_FILTER,
    CREATE_TYPE_OPTIONS, CREATE_EXT,
)
from ..thread_worker import CallWorker
from .sidebar import Sidebar
from .empty_state import CentralChat
from .preview import PreviewPanel
from .chat import ChatPanel
from .modals import SettingsDialog, ConvertDialog, OverwriteDialog, NewDocumentDialog
from .doc_creation_widgets import DocTypePicker, SuggestionCard, QuotaWarningCard, GenErrorCard
from .doc_generation import DocGenWorker
from .toast import CacheToast
from .folder_workspace import FolderWorkspace
from .folder_scan import FolderScanWorker

_MAX_RECENT = 10


def _open_path(path: str) -> None:
    if sys.platform == "win32":
        os.startfile(path)  # noqa: S606
    elif sys.platform == "darwin":
        subprocess.run(["open", path], check=False)
    else:
        subprocess.run(["xdg-open", path], check=False)


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
        self.preview.extract_text_requested.connect(self._on_extract_text)
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
        path_obj = Path(path)
        if not path_obj.exists():
            QMessageBox.warning(self, APP_NAME, f"File không tồn tại:\n{path}")
            self._remove_recent(path)
            return
        ftype = EXT_MAP.get(path_obj.suffix.lower())
        if ftype is None:
            QMessageBox.warning(self, APP_NAME, f"Định dạng chưa hỗ trợ: {path_obj.suffix}")
            return

        if fresh:
            self.chat.clear()
            self._history.clear()
            self._folder_context_name = None

        self._adopt_current_file(str(path_obj), ftype)
        verb = "Đã mở" if fresh else "Đã đính kèm"
        self.chat.add_system(f"{verb} «{path_obj.name}»")
        self.chat.input_box.setFocus()

    def _adopt_current_file(self, path: str, ftype: str):
        """Phần chung khi 1 file thật trở thành ngữ cảnh hiện tại: set state,
        đẩy vào "Gần đây", nạp preview, đặt chip theo loại file. Dùng cho cả
        mở file có sẵn (`_open_file`) và file vừa được AI tạo xong rồi lưu
        (`_on_save_generated`) — 2 luồng chỉ khác nhau ở việc có xóa hội
        thoại/lịch sử hay không."""
        self._current_path = path
        self._current_type = ftype
        self._push_recent(path)
        self._refresh_sidebar()

        self.preview.setVisible(True)
        self.preview.load_file(path, ftype, Path(path).name)
        self.chat.set_chips(CONTEXT_CHIPS.get(ftype, CONTEXT_CHIPS[None]))
        self.stack.setCurrentIndex(1)

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
        self._folder_files = [file_entry for file_entry in self._folder_files if file_entry.path != path]
        self._folder_selected = [sel_path for sel_path in self._folder_selected if sel_path != path]

    def _folder_select_all(self):
        self.sidebar.select_all_folder_files()

    def _start_folder_chat(self):
        selected = self._folder_selected or [file_entry.path for file_entry in self._folder_files]
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
        cfg["recent_files"] = [recent_path for recent_path in cfg.get("recent_files", []) if recent_path != path]
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

        # Đang mở file Word/Excel/PPT/ảnh → câu chat là lệnh sửa file (nếu AI
        # hiểu vậy). Ảnh vẫn đi qua /edit (không phải /chat) để giữ đồng nhất,
        # nhưng có kèm nội dung ảnh (xem _request_edit) nên các câu hỏi không
        # phải lệnh sửa (VD "ảnh này chụp gì") vẫn trả lời được.
        if self._current_path and self._current_type in ("word", "excel", "ppt", "image"):
            self._start_edit(text)
            return

        # Không đang sửa file mở sẵn → có thể là yêu cầu TẠO tài liệu mới
        # (7.1–7.6). Việc này xử lý cục bộ, không cần gọi server trước.
        intent = detect_create_intent(text, has_open_file=bool(self._current_path))
        if intent is not None:
            self._handle_create_intent(intent, text)
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

        # PDF đang mở → gửi kèm nội dung file để Claude đọc trực tiếp (không
        # chỉ tên file như trước) — nền tảng cho tóm tắt/trích xuất/OCR.
        # (Ảnh đi qua /edit ở nhánh phía trên, không tới đây.)
        attachment = None
        if self._current_path and self._current_type == "pdf":
            attachment = controller.build_attachment(self._current_path, self._current_type)

        self._chat_worker = CallWorker(
            api_client.chat, text, file_name or "", self._current_type or "",
            business, None, attachment,
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
            controller.request_edit, text, path, ftype, list(self._history[:-1]))
        self._edit_worker.ok.connect(self._on_edit_result)
        self._edit_worker.err.connect(self._on_chat_error)
        self._edit_worker.start()

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
        dlg = ConvertDialog(self._current_path, self._current_type or "pdf", self)
        dlg.conversion_done.connect(self._on_converted)
        dlg.exec()

    def _on_converted(self, out_path: str):
        path_obj = Path(out_path)
        if path_obj.is_dir():
            # PDF → Ảnh: nhiều file trong 1 thư mục, không phải 1 file đơn.
            count = len(list(path_obj.glob("*.png")))
            self.chat.add_system(f"✓ Đã xuất {count} ảnh vào thư mục «{path_obj.name}»")
            self.chat.add_file_card(path_obj.name, "IMG", full_path=str(path_obj))
            return
        badge = FILE_BADGE.get(EXT_MAP.get(path_obj.suffix.lower(), ""), "FILE")
        self.chat.add_system("✓ Chuyển đổi định dạng hoàn tất")
        self.chat.add_file_card(path_obj.name, badge, full_path=str(path_obj))

    # ══ Trích xuất toàn bộ văn bản PDF (cục bộ — không qua AI) ══════════════

    def _on_extract_text(self):
        if not self._current_path or self._current_type != "pdf":
            return
        try:
            content = extract_pdf(self._current_path).claude_content
        except Exception as exc:
            self.chat.add_system(f"⚠ Không trích xuất được văn bản: {exc}")
            return

        default_name = str(Path(self._current_path).with_suffix(".txt"))
        out_path, _ = QFileDialog.getSaveFileName(
            self, "Lưu văn bản đã trích xuất", default_name, "Văn bản (*.txt)")
        if not out_path:
            return
        try:
            Path(out_path).write_text(content, encoding="utf-8")
        except OSError as exc:
            self.chat.add_system(f"⚠ Không ghi được file: {exc}")
            return
        self.chat.add_system(f"✓ Đã trích xuất văn bản vào «{Path(out_path).name}»")
        self.chat.add_file_card(Path(out_path).name, "TXT", full_path=out_path)

    # ══ Tạo tài liệu mới bằng chat (7.1–7.6, EC1–EC6) ═══════════════════════

    def _handle_create_intent(self, intent: str, text: str):
        """`intent` từ `detect_create_intent()`: "word"/"excel"/"ppt" (đoán
        được định dạng), "needs_type" (chắc là tạo mới, chưa rõ định dạng),
        hoặc "ambiguous" (EC6 — câu lệnh quá mơ hồ, AI hỏi lại)."""
        self.chat.set_enabled(True)
        self.chat.input_box.setFocus()

        if intent == "ambiguous":
            card = SuggestionCard(
                "Bạn cho mình biết thêm để soạn đúng ý nhé:\n"
                "· Nội dung cụ thể là gì? (vd: doanh thu, hợp đồng, giới thiệu…)\n"
                "· Dạng văn bản (Word), bảng biểu (Excel) hay trình bày (PowerPoint)?",
                ["Báo cáo doanh thu · Excel", "Báo cáo tiến độ · Word"],
            )
            card.chip_clicked.connect(self.chat.send_text)
            self.chat.add_widget_row(card)
            return

        # EC4 — hết tác vụ AI trong tháng (gói Free): chặn trước khi tốn quota
        # thật cho bước soạn nội dung.
        if self._plan == "free" and self._quota_remaining == 0:
            card = QuotaWarningCard(
                f"Bạn đã dùng hết {self._quota_limit or 0} tác vụ AI miễn phí tháng này",
                "Nâng cấp Pro để tiếp tục tạo file ngay, hoặc chờ tới kỳ làm mới.",
            )
            card.upgrade_clicked.connect(self._open_settings)
            self.chat.add_widget_row(card)
            return

        default_key = intent if intent in ("word", "excel", "ppt") else None
        page_count = guess_page_count(text)
        self.chat.add_ai("Mình sẽ tạo giúp bạn. Xuất ra định dạng nào?")
        picker = DocTypePicker(CREATE_TYPE_OPTIONS, default_key, page_count)
        picker.confirmed.connect(
            lambda file_type: self._on_doc_type_confirmed(file_type, text, page_count))
        self.chat.add_widget_row(picker)

    def _on_doc_type_confirmed(self, file_type: str, user_text: str, page_count: int):
        """7.3 — đặt tên & chọn nơi lưu, rồi bắt đầu soạn."""
        cfg = load_config()
        default_dir = cfg.get("default_save_folder") or str(Path.home() / "Documents" / "DocAI")
        os.makedirs(default_dir, exist_ok=True)

        dlg = NewDocumentDialog(file_type, suggest_file_name(user_text), default_dir, self)
        if not dlg.exec():
            return

        cfg["default_save_folder"] = str(Path(dlg.result_path).parent)
        save_config(cfg)
        self._start_generation(file_type, user_text, page_count, dlg.result_path, dlg.open_after)

    def _start_generation(self, file_type: str, user_text: str, page_count: int,
                          final_path: str, open_after: bool):
        """7.4 — gọi AI soạn nội dung thật (qua /chat có sẵn) rồi dựng file
        thật ở luồng nền; `page_done` cập nhật checklist đang soạn."""
        self._gen_file_type = file_type
        self._gen_user_text = user_text
        self._gen_page_count = page_count
        self._gen_final_path = final_path
        self._gen_open_after = open_after

        self.chat.set_enabled(False)
        self.preview.setVisible(True)
        self.stack.setCurrentIndex(1)
        self.preview.pages.show_message("Đang soạn…")
        self.chat.add_system(f"Đang soạn «{Path(final_path).name}»…")

        business = load_config().get("business", {})
        prompt = build_creation_prompt(file_type, user_text, page_count, business)

        self._gen_worker = DocGenWorker(file_type, prompt, page_count, [], business)
        self._gen_worker.page_done.connect(self._on_gen_page_done)
        self._gen_worker.finished_ok.connect(self._on_gen_finished)
        self._gen_worker.failed.connect(self._on_gen_failed)
        self._gen_worker.start()

    def _on_gen_page_done(self, page_number: int, total: int, label: str):
        self.preview.pages.show_message(f"Đang soạn {label} ({page_number}/{total})…")

    def _on_gen_finished(self, tmp_path: str, sections, quota: dict):
        if quota.get("quota_remaining") is not None:
            self._quota_remaining = quota["quota_remaining"]
            self._plan = quota.get("plan", self._plan)
            self._update_quota_label()
        self._cache_toast.show_usage(quota.get("usage"))
        self.chat.set_enabled(True)
        self._show_preview_ready(
            tmp_path, "Đã tạo xong. Bạn xem trước rồi lưu, hoặc bảo mình chỉnh tiếp.")

    def _on_gen_failed(self, message: str, pages_done: int, partial_sections):
        """EC5 — soạn thất bại giữa chừng. `partial_sections` chỉ khác rỗng
        khi server đã trả nội dung nhưng bước dựng file cục bộ mới lỗi."""
        self.chat.set_enabled(True)
        total = self._gen_page_count if pages_done else 0
        card = GenErrorCard(message, pages_done, total, can_save_partial=bool(partial_sections))
        card.retry_clicked.connect(self._retry_generation)
        if partial_sections:
            card.save_partial_clicked.connect(lambda: self._save_partial_generation(partial_sections))
        self.chat.add_widget_row(card)
        self.preview.pages.show_message(f"Tạo file thất bại:\n{message}")

    def _retry_generation(self):
        self._start_generation(
            self._gen_file_type, self._gen_user_text, self._gen_page_count,
            self._gen_final_path, self._gen_open_after)

    def _save_partial_generation(self, partial_sections):
        try:
            fd, tmp_path = tempfile.mkstemp(
                suffix=CREATE_EXT[self._gen_file_type], prefix="docai_partial_")
            os.close(fd)
            if self._gen_file_type == "word":
                create_word(partial_sections, tmp_path)
            elif self._gen_file_type == "ppt":
                create_pptx(partial_sections, tmp_path)
            else:
                self.chat.add_system("⚠ Chưa thể lưu một phần cho file Excel.")
                return
        except Exception as exc:  # noqa: BLE001
            self.chat.add_system(f"⚠ Không lưu được phần đã có: {exc}")
            return
        self._show_preview_ready(
            tmp_path,
            f"Đã lưu tạm {len(partial_sections)} phần đã soạn được. Bạn xem trước rồi lưu.")

    def _show_preview_ready(self, tmp_path: str, message: str):
        """7.5 — xem trước file vừa dựng (chưa lưu) + nút Lưu file."""
        self._gen_tmp_path = tmp_path
        name = Path(self._gen_final_path).name
        self.preview.load_file(tmp_path, self._gen_file_type, name)
        self.chat.add_ai(message)

        labels = {"word": "Word", "excel": "Excel", "ppt": "PowerPoint"}
        save_btn = QPushButton(f"Lưu file {labels.get(self._gen_file_type, '')}")
        save_btn.setObjectName("createConfirmBtn")
        save_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        save_btn.clicked.connect(lambda: self._on_save_generated(save_btn))
        self.chat.add_widget_row(save_btn)

    def _on_save_generated(self, save_btn: QPushButton):
        """7.6 — lưu file thật vào nơi đã chọn, rồi làm việc luôn với nó
        (y hệt trạng thái đính kèm file có sẵn, không xóa hội thoại)."""
        save_btn.setEnabled(False)
        tmp_path, final_path = self._gen_tmp_path, self._gen_final_path
        try:
            Path(final_path).parent.mkdir(parents=True, exist_ok=True)
            shutil.move(tmp_path, final_path)
        except Exception as exc:  # noqa: BLE001
            self.chat.add_system(f"⚠ Không lưu được file: {exc}")
            save_btn.setEnabled(True)
            return

        self._folder_context_name = None
        self._adopt_current_file(final_path, self._gen_file_type)

        name = Path(final_path).name
        badge = FILE_BADGE.get(self._gen_file_type, "FILE")
        self.chat.add_system(
            f"✓ Đã lưu «{name}» vào {Path(final_path).parent}. Bạn muốn chỉnh gì tiếp không?")
        self.chat.add_file_card(name, badge, note="Vừa tạo")

        if self._gen_open_after:
            _open_path(final_path)

    # ══ Modal: Cài đặt ═══════════════════════════════════════════════════════

    def _open_settings(self):
        dlg = SettingsDialog(self)
        dlg.logged_out.connect(self._handle_logout)
        dlg.exec()
        self._refresh_account()

    def _handle_logout(self):
        from ...bootstrap import restart_to_login
        restart_to_login(self)

    # ══ File card (mock) ═════════════════════════════════════════════════════

    def _on_file_card(self, file_name: str):
        path_obj = Path(file_name)
        if path_obj.is_absolute() and path_obj.exists():
            _open_path(str(path_obj))
            return
        QMessageBox.information(
            self, APP_NAME,
            f"«{file_name}»\n\nFile kết quả sẽ được tạo thật khi tích hợp backend AI.")

    # ══ Dọn dẹp ══════════════════════════════════════════════════════════════

    def closeEvent(self, event):
        self.preview.cleanup()
        super().closeEvent(event)
