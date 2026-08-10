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
from ...modules.common.doc_set import process_document_set, resolve_input_files, DocSetResult
from ...modules.common.folder_ops import detect_file_mgmt_intent, create_empty_file, \
    delete_file_with_undo, restore_file, rename_file, FolderOpError
from ...modules.business.payroll import (
    detect_payroll_intent, compute_payroll_file, PayrollRunResult,
)
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
from .folder_scan import FolderScanWorker, ScannedFile
from .save_output import save_staged_file

from ...logging_config import get_logger, log_call
from ...strings import MainWindow as S

logger = get_logger(__name__)

_STAGING_DIR = str(Path(tempfile.gettempdir()) / "DocAI" / "staging")

_MAX_RECENT = 10


@log_call
def _open_path(path: str) -> None:
    if sys.platform == "win32":
        os.startfile(path)  # noqa: S606
    elif sys.platform == "darwin":
        subprocess.run(["open", path], check=False)
    else:
        subprocess.run(["xdg-open", path], check=False)


class MainWindow(QMainWindow):
    @log_call
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
        self._pending_docset_outputs: list[str] = []
        self._docset_worker: Optional[CallWorker] = None
        self._history: list[dict] = []   # {role, content} — ngữ cảnh cho AI
        self._plan = "free"
        self._quota_remaining: Optional[int] = None
        self._quota_limit: Optional[int] = None

        self._folder_path: Optional[str] = None
        self._folder_files: list = []
        self._folder_context_name: Optional[str] = None
        self._attached_files: list[str] = []   # đính kèm qua chat, cuộc trò chuyện hiện tại
        self._recent_outputs: list[str] = []   # output process_document_set(), mới nhất cuối
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

    @log_call
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
        self.sidebar.open_folder.connect(self._open_folder)
        self.sidebar.change_folder.connect(self._open_folder)
        self.sidebar.recent_deleted.connect(self._on_recent_deleted)
        self.sidebar.folder_file_removed.connect(self._on_folder_file_removed)
        body.addWidget(self.sidebar)

        # Stack: [0] chat AI trung tâm (mở app là sẵn sàng) · [1] workspace (preview + chat)
        # — Chat Panel luôn độc lập, chuyển tab Tệp/Thư mục ở sidebar không đụng
        # tới stack này; sidebar chỉ là nguồn tham chiếu file cho chat.
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

        body.addWidget(self.stack, stretch=1)
        root.addLayout(body, stretch=1)

    @log_call
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

    @log_call
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            path = event.mimeData().urls()[0].toLocalFile()
            if Path(path).suffix.lower() in EXT_MAP:
                event.acceptProposedAction()

    @log_call
    def dropEvent(self, event):
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if Path(path).suffix.lower() in EXT_MAP:
                self._open_file(path)
                return

    # ══ Mở file / luồng bắt đầu ══════════════════════════════════════════════

    @log_call
    def _pick_file(self):
        """Đính kèm file từ ô chat workspace — gắn vào cuộc trò chuyện đang có.

        Trong hội thoại chế độ Thư mục (đang xử lý nhiều file, chưa mở file
        đơn lẻ nào) — đính kèm KHÔNG chuyển sang chế độ 1-file, chỉ thêm file
        vào danh sách ứng viên cho `resolve_input_files()` (xem
        `_start_folder_task`), để không làm mất ngữ cảnh nhiều file đang có."""
        path, _ = QFileDialog.getOpenFileName(self, S.PICK_FILE_TITLE, "", FILE_DIALOG_FILTER)
        if not path:
            return
        if self._folder_context_name and not self._current_path:
            if path not in self._attached_files:
                self._attached_files.append(path)
            self.chat.add_system(S.ATTACHED_TO_CONTEXT.format(name=Path(path).name))
            return
        self._open_file(path, fresh=False)

    @log_call
    def _open_file(self, path: str, fresh: bool = True):
        path_obj = Path(path)
        if not path_obj.exists():
            QMessageBox.warning(self, APP_NAME, S.FILE_NOT_FOUND.format(path=path))
            self._remove_recent(path)
            return
        ftype = EXT_MAP.get(path_obj.suffix.lower())
        if ftype is None:
            QMessageBox.warning(self, APP_NAME, S.FORMAT_UNSUPPORTED.format(ext=path_obj.suffix))
            return

        if fresh:
            self.chat.clear()
            self._history.clear()
            self._folder_context_name = None
            self._attached_files = []
            self._recent_outputs = []

        self._adopt_current_file(str(path_obj), ftype)
        template = S.OPENED_FILE if fresh else S.ATTACHED_FILE
        self.chat.add_system(template.format(name=path_obj.name))
        self.chat.input_box.setFocus()

    @log_call
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

    @log_call
    def _start_chat(self, text: str):
        """Gõ lệnh / chip ở chat trung tâm → vào workspace, chưa cần file."""
        self._current_path = None
        self._current_type = None
        self.chat.clear()
        self.preview.setVisible(False)
        self.chat.set_chips(CONTEXT_CHIPS[None])
        self.stack.setCurrentIndex(1)
        self.chat.send_text(text)

    @log_call
    def _new_chat(self):
        """Trò chuyện mới → về chat AI trung tâm, xóa hội thoại & ngữ cảnh file."""
        self._current_path = None
        self._current_type = None
        self._folder_context_name = None
        self._attached_files = []
        self._recent_outputs = []
        self.chat.clear()
        self._history.clear()
        self.preview.setVisible(False)
        self._refresh_sidebar()
        self.stack.setCurrentIndex(0)
        self.welcome.input_box.setFocus()

    # ══ Chế độ Thư mục — mở / quét, rồi chat sẵn sàng ngay ═══════════════════
    # Chat Panel luôn là trung tâm — không có bước "tick chọn file rồi bấm Bắt
    # đầu chat"; sidebar (cây thư mục) chỉ là nguồn tham chiếu tên file cho
    # `doc_set.resolve_input_files()`. Quét xong là chat nhận lệnh được ngay.

    @log_call
    def _open_folder(self):
        path = QFileDialog.getExistingDirectory(self, S.PICK_FOLDER_TITLE, "")
        if not path:
            return

        self._folder_path = path
        self._folder_files = []

        self.sidebar.show_folder_scanning(path)

        self._folder_scan_worker = FolderScanWorker(path)
        self._folder_scan_worker.progress.connect(self._on_folder_scan_progress)
        self._folder_scan_worker.done.connect(self._on_folder_scan_done)
        self._folder_scan_worker.start()

    @log_call
    def _on_folder_scan_progress(self, done: int, total: int):
        self.sidebar.update_folder_scan_progress(done, total)

    @log_call
    def _on_folder_scan_done(self, files: list, counts: dict):
        self._folder_files = files
        folder_name = Path(self._folder_path).name or self._folder_path
        self.sidebar.show_folder_ready(folder_name, files, counts)
        self._start_folder_chat(folder_name, len(files))

    @log_call
    def _on_folder_file_removed(self, path: str):
        """Bấm xóa 1 file trong cây thư mục — chỉ gỡ khỏi danh sách tham chiếu,
        không đụng file thật."""
        self._folder_files = [f for f in self._folder_files if f.path != path]

    @log_call
    def _start_folder_chat(self, folder_name: str, file_count: int):
        self._current_path = None
        self._current_type = None
        self.chat.clear()
        self._history.clear()
        self.preview.setVisible(False)
        self.chat.set_chips(CONTEXT_CHIPS[None])

        self._attached_files = []
        self._recent_outputs = []
        self._folder_context_name = f"{folder_name} ({file_count} file)"
        self.stack.setCurrentIndex(1)
        self.chat.add_system(S.FOLDER_OPENED.format(name=folder_name, count=file_count))
        self.chat.input_box.setFocus()

    # ══ Thao tác file trong thư mục (tạo/xóa/đổi tên/mở 1 file) ═════════════
    # Theo thiết kế D:\Tools\DocAI\Plan\V4\Folder\screens_folder — nhận diện
    # cục bộ trước, không khớp mới rơi xuống process_document_set().

    @log_call
    def _handle_folder_message(self, text: str):
        folder_paths = [f.path for f in self._folder_files]
        intent = detect_file_mgmt_intent(text, folder_paths)
        if intent is None:
            self._start_folder_task(text)
            return

        if intent.ambiguous:
            self._ask_folder_mgmt_clarify(intent.kind)
            return

        if intent.kind == "create":
            self._do_create_file(intent)
        elif intent.kind == "delete":
            self._do_delete_file(intent)
        elif intent.kind == "rename":
            self._do_rename_file(intent)
        elif intent.kind == "open":
            self._do_open_folder_file(intent)

    @log_call
    def _ask_folder_mgmt_clarify(self, kind: str):
        messages = {
            "create": S.CLARIFY_CREATE,
            "delete": S.CLARIFY_DELETE,
            "rename": S.CLARIFY_RENAME,
            "open": S.CLARIFY_OPEN,
        }
        self.chat.add_ai(messages.get(kind, S.CLARIFY_GENERIC))
        self.chat.set_enabled(True)
        self.chat.input_box.setFocus()

    @log_call
    def _refresh_folder_sidebar(self):
        counts: dict[str, int] = {}
        for f in self._folder_files:
            counts[f.ext_type] = counts.get(f.ext_type, 0) + 1
        folder_name = Path(self._folder_path).name or self._folder_path
        self.sidebar.show_folder_ready(folder_name, self._folder_files, counts)

    @log_call
    def _do_create_file(self, intent):
        try:
            path = create_empty_file(self._folder_path, intent.name, intent.ftype)
        except FolderOpError as exc:
            logger.info("Create-file intent failed for %r: %s", intent.name, exc.message)
            self.chat.add_system(S.ERROR_PREFIX.format(message=exc.message))
            self.chat.set_enabled(True)
            self.chat.input_box.setFocus()
            return

        self._folder_files.append(ScannedFile(path, Path(path).name, "", intent.ftype))
        self._refresh_folder_sidebar()

        badge = FILE_BADGE.get(intent.ftype, "FILE")
        self.chat.add_system(S.CREATED_FILE.format(name=Path(path).name))
        self.chat.add_file_card(Path(path).name, badge, note=S.JUST_CREATED_NOTE, full_path=path)
        self.chat.set_enabled(True)
        self.chat.input_box.setFocus()

    @log_call
    def _do_delete_file(self, intent):
        path = intent.target_path
        try:
            trashed = delete_file_with_undo(path)
        except FolderOpError as exc:
            logger.info("Delete-file intent failed for %r: %s", path, exc.message)
            self.chat.add_system(S.ERROR_PREFIX.format(message=exc.message))
            self.chat.set_enabled(True)
            self.chat.input_box.setFocus()
            return

        name = Path(path).name
        self._folder_files = [f for f in self._folder_files if f.path != path]
        for attr in ("_attached_files", "_recent_outputs"):
            setattr(self, attr, [p for p in getattr(self, attr) if p != path])
        self._refresh_folder_sidebar()

        self.chat.add_system(S.DELETED_FILE.format(name=name))
        undo_btn = QPushButton(S.UNDO)
        undo_btn.setObjectName("undoDeleteBtn")
        undo_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        undo_btn.clicked.connect(lambda: self._undo_delete(trashed, path, undo_btn))
        self.chat.add_widget_row(undo_btn, full_width=False)
        self.chat.set_enabled(True)
        self.chat.input_box.setFocus()

    @log_call
    def _undo_delete(self, trashed_path: str, original_path: str, btn: QPushButton):
        btn.setEnabled(False)
        try:
            restore_file(trashed_path, original_path)
        except FolderOpError as exc:
            logger.info("Undo-delete failed for %r: %s", original_path, exc.message)
            self.chat.add_system(S.ERROR_PREFIX.format(message=exc.message))
            return
        ext_type = EXT_MAP.get(Path(original_path).suffix.lower(), "")
        self._folder_files.append(
            ScannedFile(original_path, Path(original_path).name, "", ext_type))
        self._refresh_folder_sidebar()
        self.chat.add_system(S.RESTORED_FILE.format(name=Path(original_path).name))

    @log_call
    def _do_rename_file(self, intent):
        try:
            new_path = rename_file(intent.target_path, intent.new_name)
        except FolderOpError as exc:
            logger.info("Rename-file intent failed for %r: %s", intent.target_path, exc.message)
            self.chat.add_system(S.ERROR_PREFIX.format(message=exc.message))
            self.chat.set_enabled(True)
            self.chat.input_box.setFocus()
            return

        old_path = intent.target_path
        for f in self._folder_files:
            if f.path == old_path:
                f.path = new_path
                f.name = Path(new_path).name
                break
        for attr in ("_attached_files", "_recent_outputs"):
            setattr(self, attr, [new_path if p == old_path else p for p in getattr(self, attr)])
        self._refresh_folder_sidebar()

        self.chat.add_system(S.RENAMED_FILE.format(name=Path(new_path).name))
        self.chat.set_enabled(True)
        self.chat.input_box.setFocus()

    @log_call
    def _do_open_folder_file(self, intent):
        path = intent.target_path
        ftype = EXT_MAP.get(Path(path).suffix.lower())
        if ftype is None:
            self.chat.add_system(S.ERROR_PREFIX.format(message=S.FORMAT_UNSUPPORTED.format(ext=Path(path).suffix)))
            self.chat.set_enabled(True)
            self.chat.input_box.setFocus()
            return

        self._adopt_current_file(path, ftype)
        self.chat.add_system(S.OPENED_FILE.format(name=Path(path).name))

        if intent.remainder:
            self._start_edit(intent.remainder)
        else:
            self.chat.set_enabled(True)
            self.chat.input_box.setFocus()

    # ══ Xử lý nhiều file (chế độ Thư mục) — so sánh/đối chiếu/gộp/trích xuất/
    # tìm kiếm/rà soát, xem `modules.common.doc_set.process_document_set()` ═══

    @log_call
    def _start_folder_task(self, text: str):
        attached = [p for p in self._attached_files if Path(p).is_file()]
        recent_outputs = [p for p in self._recent_outputs if Path(p).is_file()]
        folder_files = [f.path for f in self._folder_files if Path(f.path).is_file()]

        resolve = resolve_input_files(text, attached, recent_outputs, folder_files)
        if resolve.needs_clarification:
            self.chat.add_ai(resolve.message)
            if resolve.suggestions:
                card = SuggestionCard(S.SUGGESTIONS_IN_CONTEXT, resolve.suggestions)
                card.chip_clicked.connect(lambda name: self.chat.input_box.setText(
                    (self.chat.input_box.text() + " " + name).strip()))
                self.chat.add_widget_row(card)
            self.chat.set_enabled(True)
            self.chat.input_box.setFocus()
            return

        business = load_config().get("business", {})
        self.chat.start_ai()
        self.chat.stream_ai(S.PROCESSING)

        self._docset_worker = CallWorker(
            process_document_set, resolve.paths, text, None, list(self._history[:-1]), business)
        self._docset_worker.ok.connect(self._on_docset_result)
        self._docset_worker.err.connect(self._on_chat_error)
        self._docset_worker.start()

    @log_call
    def _on_docset_result(self, result: DocSetResult):
        if result.quota.get("quota_remaining") is not None:
            self._quota_remaining = result.quota["quota_remaining"]
            self._plan = result.quota.get("plan", self._plan)
            self._update_quota_label()

        self._pending_meta = (None, False)
        self._pending_docset_outputs = list(result.output_files)
        for out_path in result.output_files:
            if out_path not in self._recent_outputs:
                self._recent_outputs.append(out_path)
        self._pending_reply = MockReply(text=result.analysis or S.NO_ANALYSIS_CONTENT)
        self._start_stream(self._pending_reply.text)

    # ══ Sidebar / recent ═════════════════════════════════════════════════════

    @log_call
    def _push_recent(self, path: str):
        cfg = load_config()
        recent = cfg.get("recent_files", [])
        if path not in recent:
            recent = [path] + recent
        cfg["recent_files"] = recent[:_MAX_RECENT]
        save_config(cfg)

    @log_call
    def _remove_recent(self, path: str):
        cfg = load_config()
        cfg["recent_files"] = [recent_path for recent_path in cfg.get("recent_files", []) if recent_path != path]
        save_config(cfg)
        self._refresh_sidebar()

    @log_call
    def _on_recent_deleted(self, path: str):
        """Bấm xóa 1 item trong tab "Tệp" — chỉ gỡ khỏi lịch sử, không xóa file thật."""
        self._remove_recent(path)

    @log_call
    def _refresh_sidebar(self):
        recent = load_config().get("recent_files", [])
        self.sidebar.refresh(recent, self._current_path)

    # ══ Quota (nguồn: server) ════════════════════════════════════════════════

    @log_call
    def _update_quota_label(self):
        plan = "Pro" if self._plan == "pro" else "Free"
        if self._quota_remaining is None or self._quota_limit is None:
            self.quota_lbl.setText(S.PLAN_LABEL.format(plan=plan))
        else:
            self.quota_lbl.setText(S.PLAN_LABEL_WITH_QUOTA.format(
                plan=plan, remaining=self._quota_remaining, limit=self._quota_limit))
        self.sidebar.set_plan(plan, self._quota_remaining, self._quota_limit)

    @log_call
    def _refresh_account(self):
        """Lấy gói + quota từ server (nền)."""
        self._me_worker = CallWorker(api_client.me)
        self._me_worker.ok.connect(self._on_account)
        self._me_worker.err.connect(lambda _msg: None)  # im lặng nếu lỗi mạng
        self._me_worker.start()

    @log_call
    def _on_account(self, data: dict):
        self._plan = data.get("plan", "free")
        self._quota_remaining = data.get("quota_remaining")
        self._quota_limit = data.get("quota_limit")
        self._update_quota_label()

    # ══ Chat AI (gọi server, streaming khi có kết quả) ═══════════════════════

    @log_call
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

        # Đang trong hội thoại chế độ Thư mục (đã bấm "Bắt đầu trò chuyện") và
        # chưa đính kèm thêm 1 file đơn lẻ nào → định tuyến qua thao tác file
        # (tạo/xóa/đổi tên/mở 1 file) hoặc function tổng hợp nhiều file.
        if self._folder_context_name and not self._current_path:
            self._handle_folder_message(text)
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
        self.chat.stream_ai(S.PROCESSING)

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

    @log_call
    def _on_chat_result(self, data: dict):
        # Cập nhật quota từ server.
        if data.get("quota_remaining") is not None:
            self._quota_remaining = data["quota_remaining"]
            self._plan = data.get("plan", self._plan)
            self._update_quota_label()
        self._cache_toast.show_usage(data.get("usage"))

        artifact, is_edit = getattr(self, "_pending_meta", (None, False))
        self._pending_reply = MockReply(
            text=data.get("text", "") or S.NO_REPLY_CONTENT,
            artifact=artifact, is_edit=is_edit,
        )
        self._start_stream(self._pending_reply.text)

    @log_call
    def _start_stream(self, full_text: str):
        """Hiển thị dần văn bản server trả về."""
        self._stream_pos = 0
        self._stream_timer = QTimer(self)
        self._stream_timer.setInterval(16)
        self._stream_timer.timeout.connect(lambda: self._stream_tick(full_text))
        self._stream_timer.start()

    @log_call
    def _on_chat_error(self, message: str):
        self.chat.stream_ai("")
        self.chat.end_ai()
        self.chat.add_system(S.ERROR_PREFIX.format(message=message))
        self.chat.set_enabled(True)
        self.chat.input_box.setFocus()
        self._pending_reply = None

    @log_call
    def _stream_tick(self, full_text: str):
        self._stream_pos = min(len(full_text), self._stream_pos + 6)
        self.chat.stream_ai(full_text[:self._stream_pos])
        if self._stream_pos >= len(full_text):
            self._stream_timer.stop()
            self._finish_reply()

    # ══ Sửa file thật (Word / Excel) ═════════════════════════════════════════

    @log_call
    def _remember(self, role: str, content: str):
        self._history.append({"role": role, "content": content})
        del self._history[:-12]

    @log_call
    def _start_edit(self, text: str):
        path, ftype = self._current_path, self._current_type

        # "Tính lương + BHXH" trên 1 file Excel đang mở (bảng chấm công) →
        # tính TẤT ĐỊNH cục bộ (modules.business.payroll), không cho Claude
        # tự do sửa ô/công thức lương qua /edit — đúng nguyên tắc "DocAI hỗ
        # trợ, không tự kết luận số liệu" với nhóm nghiệp vụ rủi ro cao.
        # File chấm công gốc KHÔNG bị đụng tới — output là 1 file mới.
        if ftype == "excel" and detect_payroll_intent(text):
            self._start_payroll(path)
            return

        try:
            check_editable(path, ftype)
        except EditError as exc:
            logger.info("Edit request rejected for %r: %s", path, exc)
            self.chat.add_system(S.ERROR_PREFIX.format(message=exc))
            self.chat.set_enabled(True)
            self.chat.input_box.setFocus()
            return

        self.chat.start_ai()
        self.chat.stream_ai(S.PROCESSING)

        self._edit_worker = CallWorker(
            controller.request_edit, text, path, ftype, list(self._history[:-1]))
        self._edit_worker.ok.connect(self._on_edit_result)
        self._edit_worker.err.connect(self._on_chat_error)
        self._edit_worker.start()

    @log_call
    def _on_edit_result(self, data: dict):
        if data.get("quota_remaining") is not None:
            self._quota_remaining = data["quota_remaining"]
            self._plan = data.get("plan", self._plan)
            self._update_quota_label()
        self._cache_toast.show_usage(data.get("usage"))

        self._pending_meta = (None, False)
        self._pending_edit_notes = data.get("notes") or []
        self._pending_reply = MockReply(
            text=data.get("reply", "") or S.FILE_UPDATED)
        self._start_stream(self._pending_reply.text)

    @log_call
    def _start_payroll(self, path: str):
        """Tính lương + BHXH/BHYT/BHTN + TNCN cục bộ (`modules.business.payroll`)
        — không gọi AI, không tốn quota. `path` là file "bảng chấm công" đang
        mở; kết quả là 1 file "Bảng lương" mới, đi qua đúng cơ chế staging
        + file-card "Bấm để xem trước & lưu" như `process_document_set()`."""
        self.chat.start_ai()
        self.chat.stream_ai(S.CALCULATING_PAYROLL)

        self._payroll_worker = CallWorker(compute_payroll_file, path)
        self._payroll_worker.ok.connect(self._on_payroll_result)
        self._payroll_worker.err.connect(self._on_chat_error)
        self._payroll_worker.start()

    @log_call
    def _on_payroll_result(self, result: PayrollRunResult):
        self._pending_meta = (None, False)
        self._pending_docset_outputs = [result.out_path]
        if result.out_path not in self._recent_outputs:
            self._recent_outputs.append(result.out_path)

        summary = S.PAYROLL_SUMMARY.format(
            count=len(result.employees), net=result.total_net,
            nld=result.total_nld_insurance, dn=result.total_dn_cost,
            rates_version=result.rates_version)
        self._pending_reply = MockReply(text=summary)
        self._start_stream(self._pending_reply.text)

    @log_call
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
                self.chat.add_system(S.EDIT_NOTE.format(name=name, note=note))
            self.preview.load_file(self._current_path, self._current_type, name)
            return

        # Kết quả từ process_document_set() (gộp / trích xuất & gộp nguồn hỗn
        # hợp) — file mới nằm ở thư mục staging tạm, chưa lưu thật; thẻ file
        # bấm vào sẽ mở hộp thoại "Lưu file" thay vì mở trực tiếp (_on_file_card).
        outputs, self._pending_docset_outputs = self._pending_docset_outputs, []
        for out_path in outputs:
            badge = FILE_BADGE.get(EXT_MAP.get(Path(out_path).suffix.lower(), ""), "FILE")
            self.chat.add_file_card(
                Path(out_path).name, badge, note=S.STAGED_FILE_NOTE, full_path=out_path)

        if reply is None:
            return

        if reply.artifact:
            badge = FILE_BADGE.get(EXT_MAP.get(Path(reply.artifact).suffix.lower(), ""), "FILE")
            self.chat.add_file_card(reply.artifact, badge)

        if reply.is_edit and self._current_path:
            self._confirm_overwrite()

    # ══ Modal: Xác nhận ghi đè ═══════════════════════════════════════════════

    @log_call
    def _confirm_overwrite(self):
        name = Path(self._current_path).name
        dlg = OverwriteDialog(name, self)
        if dlg.exec():
            if dlg.choice == OverwriteDialog.OVERWRITE:
                self.chat.add_system(S.OVERWRITTEN.format(name=name))
            else:
                copy_name = suggest_save_name(name)
                badge = FILE_BADGE.get(self._current_type or "", "FILE")
                self.chat.add_system(S.COPY_KEPT_ORIGINAL)
                self.chat.add_file_card(copy_name, badge)

    # ══ Modal: Chuyển đổi định dạng ══════════════════════════════════════════

    @log_call
    def _open_convert(self):
        if not self._current_path:
            return
        dlg = ConvertDialog(self._current_path, self._current_type or "pdf", self)
        dlg.conversion_done.connect(self._on_converted)
        dlg.exec()

    @log_call
    def _on_converted(self, out_path: str):
        path_obj = Path(out_path)
        if path_obj.is_dir():
            # PDF → Ảnh: nhiều file trong 1 thư mục, không phải 1 file đơn.
            count = len(list(path_obj.glob("*.png")))
            self.chat.add_system(S.EXPORTED_IMAGES.format(count=count, name=path_obj.name))
            self.chat.add_file_card(path_obj.name, "IMG", full_path=str(path_obj))
            return
        badge = FILE_BADGE.get(EXT_MAP.get(path_obj.suffix.lower(), ""), "FILE")
        self.chat.add_system(S.CONVERT_DONE)
        self.chat.add_file_card(path_obj.name, badge, full_path=str(path_obj))

    # ══ Trích xuất toàn bộ văn bản PDF (cục bộ — không qua AI) ══════════════

    @log_call
    def _on_extract_text(self):
        if not self._current_path or self._current_type != "pdf":
            return
        try:
            content = extract_pdf(self._current_path).claude_content
        except Exception as exc:
            logger.exception("PDF text extraction failed for %s", self._current_path)
            self.chat.add_system(S.EXTRACT_FAILED.format(error=exc))
            return

        default_name = str(Path(self._current_path).with_suffix(".txt"))
        out_path, _ = QFileDialog.getSaveFileName(
            self, S.SAVE_EXTRACTED_TITLE, default_name, S.TEXT_FILTER)
        if not out_path:
            return
        try:
            Path(out_path).write_text(content, encoding="utf-8")
        except OSError as exc:
            logger.exception("Failed to write extracted text to %s", out_path)
            self.chat.add_system(S.WRITE_FAILED.format(error=exc))
            return
        self.chat.add_system(S.EXTRACTED_TO_FILE.format(name=Path(out_path).name))
        self.chat.add_file_card(Path(out_path).name, "TXT", full_path=out_path)

    # ══ Tạo tài liệu mới bằng chat (7.1–7.6, EC1–EC6) ═══════════════════════

    @log_call
    def _handle_create_intent(self, intent: str, text: str):
        """`intent` từ `detect_create_intent()`: "word"/"excel"/"ppt" (đoán
        được định dạng), "needs_type" (chắc là tạo mới, chưa rõ định dạng),
        hoặc "ambiguous" (EC6 — câu lệnh quá mơ hồ, AI hỏi lại)."""
        self.chat.set_enabled(True)
        self.chat.input_box.setFocus()

        if intent == "ambiguous":
            card = SuggestionCard(
                S.AMBIGUOUS_CREATE_QUESTION,
                [S.AMBIGUOUS_CREATE_CHIP_1, S.AMBIGUOUS_CREATE_CHIP_2],
            )
            card.chip_clicked.connect(self.chat.send_text)
            self.chat.add_widget_row(card)
            return

        # EC4 — hết tác vụ AI trong tháng (gói Free): chặn trước khi tốn quota
        # thật cho bước soạn nội dung.
        if self._plan == "free" and self._quota_remaining == 0:
            card = QuotaWarningCard(
                S.QUOTA_EXCEEDED.format(limit=self._quota_limit or 0),
                S.QUOTA_HINT,
            )
            card.upgrade_clicked.connect(self._open_settings)
            self.chat.add_widget_row(card)
            return

        default_key = intent if intent in ("word", "excel", "ppt") else None
        page_count = guess_page_count(text)
        self.chat.add_ai(S.ASK_FORMAT)
        picker = DocTypePicker(CREATE_TYPE_OPTIONS, default_key, page_count)
        picker.confirmed.connect(
            lambda file_type: self._on_doc_type_confirmed(file_type, text, page_count))
        self.chat.add_widget_row(picker)

    @log_call
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

    @log_call
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
        self.preview.pages.show_message(S.GENERATING_PLAIN)
        self.chat.add_system(S.GENERATING.format(name=Path(final_path).name))

        business = load_config().get("business", {})
        prompt = build_creation_prompt(file_type, user_text, page_count, business)

        self._gen_worker = DocGenWorker(file_type, prompt, page_count, [], business)
        self._gen_worker.page_done.connect(self._on_gen_page_done)
        self._gen_worker.finished_ok.connect(self._on_gen_finished)
        self._gen_worker.failed.connect(self._on_gen_failed)
        self._gen_worker.start()

    @log_call
    def _on_gen_page_done(self, page_number: int, total: int, label: str):
        self.preview.pages.show_message(S.GENERATING_PAGE.format(label=label, page=page_number, total=total))

    @log_call
    def _on_gen_finished(self, tmp_path: str, sections, quota: dict):
        if quota.get("quota_remaining") is not None:
            self._quota_remaining = quota["quota_remaining"]
            self._plan = quota.get("plan", self._plan)
            self._update_quota_label()
        self._cache_toast.show_usage(quota.get("usage"))
        self.chat.set_enabled(True)
        self._show_preview_ready(tmp_path, S.GEN_DONE)

    @log_call
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
        self.preview.pages.show_message(S.GEN_FAILED_PREVIEW.format(message=message))

    @log_call
    def _retry_generation(self):
        self._start_generation(
            self._gen_file_type, self._gen_user_text, self._gen_page_count,
            self._gen_final_path, self._gen_open_after)

    @log_call
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
                self.chat.add_system(S.CANNOT_SAVE_PARTIAL_EXCEL)
                return
        except Exception as exc:  # noqa: BLE001
            logger.exception("Failed to save partial generation (file_type=%s)",
                             self._gen_file_type)
            self.chat.add_system(S.SAVE_PARTIAL_FAILED.format(error=exc))
            return
        self._show_preview_ready(tmp_path, S.PARTIAL_SAVED.format(count=len(partial_sections)))

    @log_call
    def _show_preview_ready(self, tmp_path: str, message: str):
        """7.5 — xem trước file vừa dựng (chưa lưu) + nút Lưu file."""
        self._gen_tmp_path = tmp_path
        name = Path(self._gen_final_path).name
        self.preview.load_file(tmp_path, self._gen_file_type, name)
        self.chat.add_ai(message)

        save_btn = QPushButton(S.SAVE_FILE_TYPE.format(label=S.SAVE_LABELS.get(self._gen_file_type, "")))
        save_btn.setObjectName("createConfirmBtn")
        save_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        save_btn.clicked.connect(lambda: self._on_save_generated(save_btn))
        self.chat.add_widget_row(save_btn)

    @log_call
    def _on_save_generated(self, save_btn: QPushButton):
        """7.6 — lưu file thật vào nơi đã chọn, rồi làm việc luôn với nó
        (y hệt trạng thái đính kèm file có sẵn, không xóa hội thoại)."""
        save_btn.setEnabled(False)
        tmp_path, final_path = self._gen_tmp_path, self._gen_final_path
        try:
            Path(final_path).parent.mkdir(parents=True, exist_ok=True)
            shutil.move(tmp_path, final_path)
        except Exception as exc:  # noqa: BLE001
            logger.exception("Failed to move generated file %s -> %s", tmp_path, final_path)
            self.chat.add_system(S.SAVE_FAILED.format(error=exc))
            save_btn.setEnabled(True)
            return

        self._folder_context_name = None
        self._adopt_current_file(final_path, self._gen_file_type)

        name = Path(final_path).name
        badge = FILE_BADGE.get(self._gen_file_type, "FILE")
        self.chat.add_system(S.SAVED_AND_READY.format(name=name, dir=Path(final_path).parent))
        self.chat.add_file_card(name, badge, note=S.JUST_CREATED_NOTE)

        if self._gen_open_after:
            _open_path(final_path)

    # ══ Modal: Cài đặt ═══════════════════════════════════════════════════════

    @log_call
    def _open_settings(self):
        dlg = SettingsDialog(self)
        dlg.logged_out.connect(self._handle_logout)
        dlg.exec()
        self._refresh_account()

    @log_call
    def _handle_logout(self):
        from ...bootstrap import restart_to_login
        restart_to_login(self)

    # ══ File card (mock) ═════════════════════════════════════════════════════

    @log_call
    def _on_file_card(self, file_name: str):
        path_obj = Path(file_name)
        if path_obj.is_absolute() and path_obj.exists():
            if self._is_staged_output(path_obj):
                self._save_and_open_staged(path_obj)
                return
            self._open_file(str(path_obj))
            return
        QMessageBox.information(self, APP_NAME, S.MOCK_FILE_RESULT.format(name=file_name))

    @log_call
    def _is_staged_output(self, path_obj: Path) -> bool:
        """File kết quả từ process_document_set() (gộp / trích xuất & gộp nguồn
        hỗn hợp) — chưa được lưu thật, còn nằm ở thư mục staging tạm."""
        try:
            path_obj.relative_to(_STAGING_DIR)
            return True
        except ValueError:
            return False

    @log_call
    def _save_and_open_staged(self, staged_path: Path):
        """Lưu file kết quả (gộp / trích xuất & gộp nguồn hỗn hợp) ra vị trí
        người dùng chọn, rồi tiếp tục làm việc với nó ngay trong cuộc trò
        chuyện hiện tại — giữ nguyên lịch sử chat, không xóa gì cả (giống
        `_on_save_generated`, khớp thiết kế 3.2b "mở như 1 file bình thường")."""
        saved_path = save_staged_file(self, str(staged_path), staged_path.name)
        if not saved_path:
            return
        self._folder_context_name = None
        self._adopt_current_file(saved_path, EXT_MAP.get(Path(saved_path).suffix.lower(), ""))
        self.chat.add_system(S.SAVED_CONTINUE.format(name=Path(saved_path).name))

    # ══ Dọn dẹp ══════════════════════════════════════════════════════════════

    @log_call
    def closeEvent(self, event):
        self.preview.cleanup()
        super().closeEvent(event)
