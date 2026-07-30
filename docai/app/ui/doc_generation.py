"""Worker sinh tài liệu mới bằng chat AI (7.4) — gọi `/chat` (server thật)
lấy nội dung, tách trang/sheet, rồi dựng file .docx/.xlsx/.pptx thật ở luồng
nền, phát tiến trình từng trang cho UI (checklist 7.4)."""
import os
import re
import tempfile

from PySide6.QtCore import QThread, Signal

from ...account import api_client
from ...account.api_client import ApiError
from ...modules.common.creators import (
    SectionSpec, parse_sections, create_word, create_pptx, create_excel_from_text,
)

_PAGE_PACE_MS = 220
_SHEET_RE = re.compile(r'^-{3}\s*Sheet:\s*(.+?)\s*-{3}\s*$', re.IGNORECASE | re.MULTILINE)
_SUFFIX = {"word": ".docx", "excel": ".xlsx", "ppt": ".pptx"}


class DocGenWorker(QThread):
    """page_done(i, total, label) trong lúc dựng; finished_ok(tmp_path,
    sections, quota_data) khi xong; failed(message, pages_done,
    partial_sections) khi lỗi — `partial_sections` chỉ khác rỗng khi nội
    dung đã lấy được từ server nhưng bước dựng file cục bộ thất bại (khi đó
    có thể lưu tạm phần đã có — EC5)."""

    page_done = Signal(int, int, str)
    finished_ok = Signal(str, object, dict)
    failed = Signal(str, int, object)

    def __init__(self, file_type: str, prompt: str, page_count: int,
                 history: list | None = None, business: dict | None = None, parent=None):
        super().__init__(parent)
        self._file_type = file_type
        self._prompt = prompt
        self._page_count = page_count
        self._history = history or []
        self._business = business or {}

    def run(self):
        try:
            data = api_client.chat(self._prompt, "", "", self._business, self._history)
        except ApiError as exc:
            self.failed.emit(exc.message, 0, [])
            return
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(f"Lỗi không xác định: {exc}", 0, [])
            return

        raw_text = (data.get("text") or "").strip()
        if not raw_text:
            self.failed.emit("AI không trả về nội dung để tạo file.", 0, [])
            return

        if self._file_type == "excel":
            names = [match.group(1).strip() for match in _SHEET_RE.finditer(raw_text)] or ["Bảng tính"]
            sections = [SectionSpec(heading=name, body="") for name in names]
        else:
            sections = parse_sections(raw_text, self._page_count)

        if not sections:
            self.failed.emit("AI không trả về nội dung để tạo file.", 0, [])
            return

        total = len(sections)
        for page_number, sec in enumerate(sections, start=1):
            self.msleep(_PAGE_PACE_MS)
            self.page_done.emit(page_number, total, sec.heading)

        try:
            fd, tmp_path = tempfile.mkstemp(suffix=_SUFFIX[self._file_type], prefix="docai_new_")
            os.close(fd)
            if self._file_type == "word":
                create_word(sections, tmp_path)
            elif self._file_type == "ppt":
                create_pptx(sections, tmp_path)
            else:
                create_excel_from_text(raw_text, tmp_path)
        except Exception as exc:  # noqa: BLE001 — lỗi cục bộ, nội dung đã soạn vẫn còn
            self.failed.emit(f"Không dựng được file: {exc}", total, sections)
            return

        quota = {
            "quota_remaining": data.get("quota_remaining"),
            "plan": data.get("plan"),
            "usage": data.get("usage"),
        }
        self.finished_ok.emit(tmp_path, sections, quota)
