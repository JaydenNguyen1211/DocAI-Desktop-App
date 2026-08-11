"""Secondary modals: Settings & account · Convert format · Confirm overwrite ·
Create new document."""
import os
import re
from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QLineEdit,
    QComboBox, QProgressBar, QFrame, QStackedWidget, QWidget, QFormLayout,
    QCheckBox, QFileDialog, QButtonGroup,
)

from ...account import api_client
from ...account.config import load_config, save_config
from ...modules.common.converters import convert_file
from ...modules.common.filenames import FORBIDDEN_FILENAME_CHARS
from ..constants import CONVERT_TARGETS, CONVERT_EXT, CREATE_EXT, CREATE_TITLE
from ..thread_worker import CallWorker

from ...logging_config import get_logger, log_call
from ...strings import (
    SettingsDialog as S_Settings, ConvertDialog as S_Convert,
    OverwriteDialog as S_Overwrite, NewDocumentDialog as S_NewDoc,
)

logger = get_logger(__name__)


@log_call
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


# ══ 4.4 Settings & account ════════════════════════════════════════════════

class SettingsDialog(QDialog):
    logged_out = Signal()

    @log_call
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(S_Settings.TITLE)
        self.setModal(True)
        self.setFixedSize(640, 560)
        self.setStyleSheet("QDialog { background: white; }")

        self._me_worker: CallWorker | None = None
        self._biz_worker: CallWorker | None = None
        cfg = load_config()
        lay = QVBoxLayout(self)
        lay.setContentsMargins(28, 24, 28, 20)
        lay.setSpacing(14)
        lay.addLayout(_modal_header(self, S_Settings.TITLE))

        # ── Tab buttons ───────────────────────────────────────────────────────
        tabs_row = QHBoxLayout()
        tabs_row.setSpacing(8)
        self._tab_btns: list[QPushButton] = []
        for tab_index, name in enumerate(
                [S_Settings.TAB_ACCOUNT, S_Settings.TAB_BUSINESS, S_Settings.TAB_LIMITS]):
            btn = QPushButton(name)
            btn.setProperty("class", "tabBtn")
            btn.setCheckable(True)
            btn.setChecked(tab_index == 0)
            btn.clicked.connect(lambda _, idx=tab_index: self._switch_tab(idx))
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
        save_btn = QPushButton(S_Settings.SAVE)
        save_btn.setProperty("class", "primaryBtn")
        save_btn.clicked.connect(self._save)
        save_row.addWidget(save_btn)
        lay.addLayout(save_row)

        note = QLabel(S_Settings.FOOTNOTE)
        note.setObjectName("footnote")
        lay.addWidget(note)

        self._load_account()

    @log_call
    def _switch_tab(self, idx: int):
        for btn_index, btn in enumerate(self._tab_btns):
            btn.setChecked(btn_index == idx)
        self._stack.setCurrentIndex(idx)

    # ── Account tab ──────────────────────────────────────────────────────────

    @log_call
    def _build_account_tab(self, cfg: dict) -> QWidget:
        tab = QWidget()
        lay = QVBoxLayout(tab)
        lay.setContentsMargins(0, 10, 0, 0)
        lay.setSpacing(12)

        self._acc_vals: dict[str, QLabel] = {}
        for key, label, initial in [
            ("email", S_Settings.FIELD_ACCOUNT, api_client.current_email() or "—"),
            ("plan", S_Settings.FIELD_PLAN, S_Settings.LOADING),
            ("quota", S_Settings.FIELD_QUOTA, S_Settings.LOADING),
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
        logout_btn = QPushButton(S_Settings.LOGOUT)
        logout_btn.setProperty("class", "secondaryBtn")
        logout_btn.clicked.connect(self._logout)
        logout_row.addWidget(logout_btn)
        lay.addLayout(logout_row)
        return tab

    @log_call
    def _load_account(self):
        self._me_worker = CallWorker(api_client.me)
        self._me_worker.ok.connect(self._on_account)
        self._me_worker.err.connect(lambda msg: self._acc_vals["plan"].setText(msg))
        self._me_worker.start()

    @log_call
    def _on_account(self, data: dict):
        plan = S_Settings.PLAN_PRO if data.get("plan") == "pro" else S_Settings.PLAN_FREE
        self._acc_vals["email"].setText(data.get("email") or "—")
        self._acc_vals["plan"].setText(plan)
        self._acc_vals["quota"].setText(
            S_Settings.QUOTA_VALUE.format(
                remaining=data.get("quota_remaining", 0), limit=data.get("quota_limit", 0)))
        # Sync business info from the server into the input fields, if present.
        biz = data.get("business") or {}
        for biz_key, edit in getattr(self, "biz_inputs", {}).items():
            if biz.get(biz_key) and not edit.text().strip():
                edit.setText(str(biz[biz_key]))

    @log_call
    def _logout(self):
        api_client.logout()
        self.logged_out.emit()
        self.accept()

    # ── Business tab ─────────────────────────────────────────────────────────

    @log_call
    def _build_business_tab(self, cfg: dict) -> QWidget:
        tab = QWidget()
        form = QFormLayout(tab)
        form.setContentsMargins(0, 10, 0, 0)
        form.setVerticalSpacing(12)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignLeft)

        biz = cfg.get("business", {})
        self.biz_inputs: dict[str, QLineEdit] = {}
        for key, label in [
            ("company_name", S_Settings.FIELD_COMPANY_NAME),
            ("tax_code", S_Settings.FIELD_TAX_CODE),
            ("address", S_Settings.FIELD_ADDRESS),
            ("representative", S_Settings.FIELD_REPRESENTATIVE),
        ]:
            lbl = QLabel(label)
            lbl.setObjectName("fieldLabel")
            edit = QLineEdit(biz.get(key, ""))
            self.biz_inputs[key] = edit
            form.addRow(lbl, edit)
        return tab

    # ── Limits tab ────────────────────────────────────────────────────────────

    @log_call
    def _build_limits_tab(self, cfg: dict) -> QWidget:
        tab = QWidget()
        lay = QVBoxLayout(tab)
        lay.setContentsMargins(0, 10, 0, 0)
        lbl = QLabel(S_Settings.LIMITS_PLACEHOLDER)
        lbl.setObjectName("fieldLabel")
        lay.addWidget(lbl)
        lay.addStretch()
        return tab

    @log_call
    def _save(self):
        business = {field_key: field_edit.text().strip() for field_key, field_edit in self.biz_inputs.items()}
        cfg = load_config()
        cfg["business"] = business   # save locally so chat can use it right away
        save_config(cfg)
        # Sync to the server (background); don't block closing the dialog.
        self._biz_worker = CallWorker(api_client.update_business, business)
        self._biz_worker.err.connect(lambda _msg: None)
        self._biz_worker.start()
        self.accept()


# ══ 4.5 Convert format ═══════════════════════════════════════════════════

class ConvertDialog(QDialog):
    conversion_done = Signal(str)   # real result path (file or folder — see PDF→Image)

    @log_call
    def __init__(self, path: str, file_type: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle(S_Convert.TITLE)
        self.setModal(True)
        self.setFixedSize(500, 350 if file_type == "image" else 300)
        self.setStyleSheet("QDialog { background: white; }")
        self._path = path
        self._file_name = Path(path).name
        self._file_type = file_type
        self._extra_paths: list[str] = []
        self._worker: CallWorker | None = None

        lay = QVBoxLayout(self)
        lay.setContentsMargins(28, 24, 28, 20)
        lay.setSpacing(12)
        lay.addLayout(_modal_header(self, S_Convert.TITLE))

        src = QLabel(self._file_name)
        src.setObjectName("fieldLabel")
        lay.addWidget(src)
        lay.addSpacing(6)

        target_lbl = QLabel(S_Convert.TARGET_FORMAT_LABEL)
        target_lbl.setObjectName("fieldLabel")
        lay.addWidget(target_lbl)

        self.target_combo = QComboBox()
        self.target_combo.addItems(CONVERT_TARGETS.get(file_type, ["PDF (.pdf)"]))
        lay.addWidget(self.target_combo)

        # Merge multiple images into 1 multi-page file (PDF/Word) — only
        # applies when the source is an image. Extra images are processed
        # AFTER the original, in the order they were picked.
        if file_type == "image":
            extra_row = QHBoxLayout()
            self.extra_lbl = QLabel(S_Convert.NO_EXTRA_IMAGES)
            self.extra_lbl.setObjectName("fieldLabel")
            extra_row.addWidget(self.extra_lbl)
            extra_row.addStretch()
            add_extra_btn = QPushButton(S_Convert.ADD_EXTRA_IMAGE)
            add_extra_btn.setProperty("class", "secondaryBtn")
            add_extra_btn.clicked.connect(self._pick_extra_images)
            extra_row.addWidget(add_extra_btn)
            lay.addLayout(extra_row)

        self.status_lbl = QLabel("")
        self.status_lbl.setObjectName("fieldLabel")
        self.status_lbl.setVisible(False)
        lay.addWidget(self.status_lbl)

        # There's no real % for COM/AI-based conversion — use an indeterminate
        # (marquee) mode instead of the fake progress bar from before.
        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.setFixedHeight(8)
        self.progress.setVisible(False)
        lay.addWidget(self.progress)
        lay.addStretch()

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.cancel_btn = QPushButton(S_Convert.CANCEL)
        self.cancel_btn.setProperty("class", "secondaryBtn")
        self.cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(self.cancel_btn)
        self.convert_btn = QPushButton(S_Convert.CONVERT)
        self.convert_btn.setProperty("class", "primaryBtn")
        self.convert_btn.clicked.connect(self._start)
        btn_row.addWidget(self.convert_btn)
        lay.addLayout(btn_row)

    @log_call
    def _target_ext(self) -> str:
        return CONVERT_EXT[self.target_combo.currentText()]

    @log_call
    def _is_searchable(self) -> bool:
        return self.target_combo.currentText() == "PDF - có thể tìm kiếm (.pdf)"

    @log_call
    def _pick_extra_images(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, S_Convert.PICK_EXTRA_TITLE,
            str(Path(self._path).parent),
            S_Convert.IMAGE_FILTER,
        )
        already = {self._path, *self._extra_paths}
        for picked_path in paths:
            if picked_path not in already:
                self._extra_paths.append(picked_path)
                already.add(picked_path)
        extra_count = len(self._extra_paths)
        self.extra_lbl.setText(
            S_Convert.NO_EXTRA_IMAGES if extra_count == 0
            else S_Convert.EXTRA_IMAGES_ADDED.format(count=extra_count))

    @log_call
    def _compute_out_path(self, target_ext: str) -> str:
        src = Path(self._path)
        if self._file_type == "pdf" and target_ext == ".png":
            return str(src.with_name(f"{src.stem}_anh"))
        return str(src.with_suffix(target_ext))

    @log_call
    def _start(self):
        self.convert_btn.setEnabled(False)
        self.cancel_btn.setEnabled(False)
        self.target_combo.setEnabled(False)
        self.progress.setVisible(True)
        self.status_lbl.setText(S_Convert.CONVERTING)
        self.status_lbl.setVisible(True)

        target_ext = self._target_ext()
        out_path = self._compute_out_path(target_ext)
        self._worker = CallWorker(
            convert_file, self._path, self._file_type, target_ext, out_path,
            self._extra_paths, self._is_searchable())
        self._worker.ok.connect(self._on_convert_ok)
        self._worker.err.connect(self._on_convert_err)
        self._worker.start()

    @log_call
    def _on_convert_ok(self, out_path: str):
        self.conversion_done.emit(out_path)
        self.accept()

    @log_call
    def _on_convert_err(self, msg: str):
        self.progress.setVisible(False)
        self.status_lbl.setText(S_Convert.ERROR_PREFIX.format(message=msg))
        self.convert_btn.setEnabled(True)
        self.cancel_btn.setEnabled(True)
        self.target_combo.setEnabled(True)


# ══ 4.6 Confirm file overwrite ═══════════════════════════════════════════

class OverwriteDialog(QDialog):
    SAVE_COPY = 1
    OVERWRITE = 2

    @log_call
    def __init__(self, file_name: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle(S_Overwrite.TITLE)
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
        title = QLabel(S_Overwrite.HEADING)
        title.setObjectName("modalTitle")
        head.addWidget(title)
        head.addStretch()
        lay.addLayout(head)

        msg = QLabel(S_Overwrite.MESSAGE.format(file_name=file_name))
        msg.setObjectName("fieldLabel")
        msg.setWordWrap(True)
        lay.addWidget(msg)
        lay.addStretch()

        btn_row = QHBoxLayout()
        btn_row.setSpacing(12)
        save_copy = QPushButton(S_Overwrite.SAVE_COPY)
        save_copy.setProperty("class", "secondaryBtn")
        save_copy.clicked.connect(self._pick_copy)
        btn_row.addWidget(save_copy, stretch=1)
        overwrite = QPushButton(S_Overwrite.OVERWRITE)
        overwrite.setProperty("class", "primaryBtn")
        overwrite.clicked.connect(self._pick_overwrite)
        btn_row.addWidget(overwrite, stretch=1)
        lay.addLayout(btn_row)

    @log_call
    def _pick_copy(self):
        self.choice = self.SAVE_COPY
        self.accept()

    @log_call
    def _pick_overwrite(self):
        self.choice = self.OVERWRITE
        self.accept()


# ══ 7.3 Create new document via chat — name & save location ═══════════════

class NewDocumentDialog(QDialog):
    """Set the name + save location before the AI writes the new file.
    Handled inline: EC3 (empty name/forbidden characters) blocks the Create
    File button; EC1 (duplicate name) shows a rename/overwrite choice; EC2
    (invalid save location) switches the dialog to a full-screen error
    state. After a successful `exec()`: `.result_path` (the final path)
    and `.open_after` (whether to open the file right away)."""

    @log_call
    def __init__(self, file_type: str, suggested_name: str, default_dir: str, parent=None):
        super().__init__(parent)
        self._file_type = file_type
        self._ext = CREATE_EXT.get(file_type, "")
        self._save_dir = default_dir
        self._dup_choice: str | None = None
        self.result_path: str | None = None
        self.open_after: bool = False

        title = CREATE_TITLE.get(file_type, S_NewDoc.DEFAULT_TITLE)
        self.setWindowTitle(title)
        self.setModal(True)
        self.setFixedSize(460, 460)
        self.setStyleSheet("QDialog { background: white; }")

        outer = QVBoxLayout(self)
        outer.setContentsMargins(26, 22, 26, 22)
        outer.setSpacing(10)
        outer.addLayout(_modal_header(self, title))

        self._stack = QStackedWidget()
        outer.addWidget(self._stack, stretch=1)
        self._stack.addWidget(self._build_form_page(suggested_name))
        self._stack.addWidget(self._build_path_error_page())

        self._validate_name()

    # ── Main page: name + save location ─────────────────────────────────────

    @log_call
    def _build_form_page(self, suggested_name: str) -> QWidget:
        page = QWidget()
        lay = QVBoxLayout(page)
        lay.setContentsMargins(0, 4, 0, 0)
        lay.setSpacing(4)

        name_lbl = QLabel(S_NewDoc.FILE_NAME_LABEL)
        name_lbl.setObjectName("fieldLabel")
        lay.addWidget(name_lbl)

        self._name_frame = QFrame()
        self._name_frame.setFixedHeight(44)
        nf_lay = QHBoxLayout(self._name_frame)
        nf_lay.setContentsMargins(13, 0, 12, 0)
        self.name_edit = QLineEdit(suggested_name)
        self.name_edit.setStyleSheet("border: none; font-size: 13.5px; font-weight: 500;")
        self.name_edit.textChanged.connect(self._on_name_changed)
        nf_lay.addWidget(self.name_edit, stretch=1)
        ext_lbl = QLabel(self._ext)
        ext_lbl.setStyleSheet(
            "color:#A8A49C; font-family:Consolas,monospace; font-size:13px; background:transparent;")
        nf_lay.addWidget(ext_lbl)
        lay.addWidget(self._name_frame)

        self._name_error = QLabel("")
        self._name_error.setStyleSheet("color:#C0392B; font-size:11.5px; background:transparent;")
        self._name_error.setWordWrap(True)
        self._name_error.setVisible(False)
        lay.addWidget(self._name_error)

        lay.addSpacing(6)
        loc_lbl = QLabel(S_NewDoc.SAVE_LOCATION_LABEL)
        loc_lbl.setObjectName("fieldLabel")
        lay.addWidget(loc_lbl)

        loc_row = QFrame()
        loc_row.setFixedHeight(44)
        loc_row.setStyleSheet("QFrame { border:1px solid #D8D5CF; border-radius:9px; }")
        lr_lay = QHBoxLayout(loc_row)
        lr_lay.setContentsMargins(13, 0, 6, 0)
        self._path_lbl = QLabel(self._save_dir)
        self._path_lbl.setStyleSheet(
            "font-family:Consolas,monospace; font-size:11.5px; color:#4A4842; background:transparent;")
        lr_lay.addWidget(self._path_lbl, stretch=1)
        change_btn = QPushButton(S_NewDoc.CHANGE)
        change_btn.setProperty("class", "secondaryBtn")
        change_btn.setFixedHeight(32)
        change_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        change_btn.clicked.connect(self._pick_dir)
        lr_lay.addWidget(change_btn)
        lay.addWidget(loc_row)

        lay.addSpacing(6)
        self.open_check = QCheckBox(S_NewDoc.OPEN_AFTER_CREATE)
        self.open_check.setChecked(True)
        lay.addWidget(self.open_check)

        # EC1 — duplicate name: hidden by default, shown when "Create File" is
        # clicked and a duplicate is found.
        self._dup_host = QWidget()
        dup_lay = QVBoxLayout(self._dup_host)
        dup_lay.setContentsMargins(0, 10, 0, 0)
        dup_lay.setSpacing(7)
        dup_warn = QLabel(S_NewDoc.DUPLICATE_WARNING)
        dup_warn.setStyleSheet("color:#C0392B; font-size:11.5px; background:transparent;")
        dup_lay.addWidget(dup_warn)
        self._dup_group = QButtonGroup(self)
        self._dup_group.setExclusive(True)
        self._rename_btn = QPushButton()
        self._rename_btn.setCheckable(True)
        self._rename_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._rename_btn.setStyleSheet(self._dup_option_qss())
        self._rename_btn.clicked.connect(lambda: self._pick_dup("rename"))
        self._overwrite_btn = QPushButton(S_NewDoc.OVERWRITE_OLD)
        self._overwrite_btn.setCheckable(True)
        self._overwrite_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._overwrite_btn.setStyleSheet(self._dup_option_qss())
        self._overwrite_btn.clicked.connect(lambda: self._pick_dup("overwrite"))
        self._dup_group.addButton(self._rename_btn)
        self._dup_group.addButton(self._overwrite_btn)
        dup_lay.addWidget(self._rename_btn)
        dup_lay.addWidget(self._overwrite_btn)
        self._dup_host.setVisible(False)
        lay.addWidget(self._dup_host)

        lay.addStretch()

        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)
        btn_row.addStretch()
        cancel = QPushButton(S_NewDoc.CANCEL)
        cancel.setProperty("class", "secondaryBtn")
        cancel.setCursor(Qt.CursorShape.PointingHandCursor)
        cancel.clicked.connect(self.reject)
        btn_row.addWidget(cancel)
        self.create_btn = QPushButton(S_NewDoc.CREATE_FILE)
        self.create_btn.setProperty("class", "primaryBtn")
        self.create_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.create_btn.clicked.connect(self._on_create)
        btn_row.addWidget(self.create_btn)
        lay.addLayout(btn_row)
        return page

    @staticmethod
    @log_call
    def _dup_option_qss() -> str:
        return (
            "QPushButton { text-align:left; border:1px solid #D8D5CF; border-radius:9px;"
            " padding:0 12px; font-size:12.5px; font-weight:500; height:38px; background:white; }"
            "QPushButton:checked { border-color:#2A6FDB; background:#EAF1FC; }"
        )

    # ── Save-location error page (EC2) ──────────────────────────────────────

    @log_call
    def _build_path_error_page(self) -> QWidget:
        page = QWidget()
        lay = QVBoxLayout(page)
        lay.setContentsMargins(0, 24, 0, 0)
        lay.setSpacing(16)

        head = QHBoxLayout()
        head.setSpacing(13)
        icon = QLabel("!")
        icon.setObjectName("dangerIcon")
        icon.setFixedSize(40, 32)
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        head.addWidget(icon)
        text_col = QVBoxLayout()
        text_col.setSpacing(4)
        title = QLabel(S_NewDoc.PATH_ERROR_TITLE)
        title.setStyleSheet("font-size:16.5px; font-weight:700;")
        text_col.addWidget(title)
        self._path_error_msg = QLabel("")
        self._path_error_msg.setObjectName("fieldLabel")
        self._path_error_msg.setWordWrap(True)
        text_col.addWidget(self._path_error_msg)
        head.addLayout(text_col, stretch=1)
        lay.addLayout(head)
        lay.addStretch()

        choose_btn = QPushButton(S_NewDoc.CHOOSE_OTHER_LOCATION)
        choose_btn.setProperty("class", "primaryBtn")
        choose_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        choose_btn.clicked.connect(self._pick_dir_from_error)
        lay.addWidget(choose_btn)
        retry_btn = QPushButton(S_NewDoc.RETRY)
        retry_btn.setProperty("class", "secondaryBtn")
        retry_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        retry_btn.clicked.connect(lambda: self._stack.setCurrentIndex(0))
        lay.addWidget(retry_btn)
        return page

    # ── Validate the name (EC3) ─────────────────────────────────────────────

    @log_call
    def _on_name_changed(self, _text: str):
        self._dup_choice = None
        self._dup_host.setVisible(False)
        self._validate_name()

    @log_call
    def _validate_name(self):
        name = self.name_edit.text().strip()
        forbidden = sorted(set(re.findall(FORBIDDEN_FILENAME_CHARS, name)))
        if not name:
            self._set_name_error(S_NewDoc.NAME_EMPTY)
        elif forbidden:
            self._set_name_error(S_NewDoc.NAME_FORBIDDEN.format(chars=" ".join(forbidden)))
        else:
            self._clear_name_error()

    @log_call
    def _set_name_error(self, msg: str):
        self._name_frame.setStyleSheet(
            "QFrame { border:1.5px solid #C0392B; border-radius:9px; }")
        self._name_error.setText(f"⚠ {msg}")
        self._name_error.setVisible(True)
        self.create_btn.setEnabled(False)

    @log_call
    def _clear_name_error(self):
        self._name_frame.setStyleSheet(
            "QFrame { border:1px solid #D8D5CF; border-radius:9px; }")
        self._name_error.setVisible(False)
        self.create_btn.setEnabled(True)

    # ── Pick / change the save location ─────────────────────────────────────

    @log_call
    def _pick_dir(self):
        path = QFileDialog.getExistingDirectory(self, S_NewDoc.PICK_SAVE_DIR_TITLE, self._save_dir)
        if path:
            self._save_dir = path
            self._path_lbl.setText(path)
            self._dup_choice = None
            self._dup_host.setVisible(False)

    @log_call
    def _pick_dir_from_error(self):
        path = QFileDialog.getExistingDirectory(self, S_NewDoc.PICK_SAVE_DIR_TITLE, str(Path.home()))
        if path:
            self._save_dir = path
            self._path_lbl.setText(path)
        self._stack.setCurrentIndex(0)

    # ── Create the file ──────────────────────────────────────────────────────

    @log_call
    def _pick_dup(self, choice: str):
        self._dup_choice = choice
        self.create_btn.setEnabled(True)

    @log_call
    def _on_create(self):
        name = self.name_edit.text().strip()
        if not name or re.search(FORBIDDEN_FILENAME_CHARS, name):
            return

        # EC2 — the save folder no longer exists / no write permission.
        if not Path(self._save_dir).is_dir() or not os.access(self._save_dir, os.W_OK):
            self._path_error_msg.setText(S_NewDoc.PATH_ERROR_MSG.format(dir=self._save_dir))
            self._stack.setCurrentIndex(1)
            return

        target = Path(self._save_dir) / f"{name}{self._ext}"

        # EC1 — duplicate name: need the user to choose how to handle it before continuing.
        if target.exists() and self._dup_choice is None:
            self._rename_btn.setText(S_NewDoc.RENAME_TO.format(name=f"{name} (1){self._ext}"))
            self._dup_host.setVisible(True)
            self.create_btn.setEnabled(False)
            return

        if self._dup_choice == "rename":
            target = Path(self._save_dir) / f"{name} (1){self._ext}"

        self.result_path = str(target)
        self.open_after = self.open_check.isChecked()
        self.accept()
