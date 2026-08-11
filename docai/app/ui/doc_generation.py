"""Worker that generates a new document via AI chat (7.4) — calls `/chat`
(the real server) to get content, splits it into pages/sheets, then builds
the real .docx/.xlsx/.pptx file on a background thread, emitting per-page
progress for the UI (7.4 checklist)."""
import os
import re
import tempfile

from PySide6.QtCore import QThread, Signal

from ...account import api_client
from ...account.api_client import ApiError
from ...modules.common.creators import (
    SectionSpec, parse_sections, create_word, create_pptx, create_excel_from_text,
)

from ...logging_config import get_logger, log_call
from ...strings import DocGeneration as S

logger = get_logger(__name__)

_PAGE_PACE_MS = 220
_SHEET_RE = re.compile(r'^-{3}\s*Sheet:\s*(.+?)\s*-{3}\s*$', re.IGNORECASE | re.MULTILINE)
_SUFFIX = {"word": ".docx", "excel": ".xlsx", "ppt": ".pptx"}


class DocGenWorker(QThread):
    """page_done(i, total, label) while building; finished_ok(tmp_path,
    sections, quota_data) when done; failed(message, pages_done,
    partial_sections) on error — `partial_sections` is non-empty only when
    content was already fetched from the server but the local file-building
    step failed (in that case the partial content can still be saved — EC5)."""

    page_done = Signal(int, int, str)
    finished_ok = Signal(str, object, dict)
    failed = Signal(str, int, object)

    @log_call
    def __init__(self, file_type: str, prompt: str, page_count: int,
                 history: list | None = None, business: dict | None = None, parent=None):
        super().__init__(parent)
        self._file_type = file_type
        self._prompt = prompt
        self._page_count = page_count
        self._history = history or []
        self._business = business or {}

    @log_call
    def run(self):
        try:
            data = api_client.chat(self._prompt, "", "", self._business, self._history)
        except ApiError as exc:
            logger.warning("Document generation chat call failed: %s (code=%s)",
                           exc.message, exc.code)
            self.failed.emit(exc.message, 0, [])
            return
        except Exception as exc:  # noqa: BLE001
            logger.exception("Document generation chat call failed unexpectedly")
            self.failed.emit(S.UNKNOWN_ERROR.format(error=exc), 0, [])
            return

        raw_text = (data.get("text") or "").strip()
        if not raw_text:
            self.failed.emit(S.NO_CONTENT, 0, [])
            return

        if self._file_type == "excel":
            names = [match.group(1).strip() for match in _SHEET_RE.finditer(raw_text)] or [S.DEFAULT_SHEET_NAME]
            sections = [SectionSpec(heading=name, body="") for name in names]
        else:
            sections = parse_sections(raw_text, self._page_count)

        if not sections:
            self.failed.emit(S.NO_CONTENT, 0, [])
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
        except Exception as exc:  # noqa: BLE001 — local error, generated content is still around
            logger.exception("Failed to build generated %s file locally", self._file_type)
            self.failed.emit(S.BUILD_FAILED.format(error=exc), total, sections)
            return

        quota = {
            "quota_remaining": data.get("quota_remaining"),
            "plan": data.get("plan"),
            "usage": data.get("usage"),
        }
        self.finished_ok.emit(tmp_path, sections, quota)
