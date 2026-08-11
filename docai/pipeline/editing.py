"""Apply the AI's edit commands to a Word/Excel/PowerPoint/image file in
place (overwrites the original file).

The server (Claude) decides the list of edit commands; this module just
dispatches to the right engine in `modules/` based on file type.
  · Word    — full operation set in `modules.word.ops` (add/remove paragraph,
              delete page, insert at a specific position, formatting, find &
              replace…).
  · Excel   — full operation set in `modules.excel.ops` (cell/range,
              formatting, tables, sort, filter, charts, sheets, freeze
              pane…).
  · PPT     — operation set in `modules.pptx.ops` (text/formatting/lists,
              add-delete-move slides, insert images).
  · Image   — operation set in `modules.image.ops` (rotate/crop/resize/
              compress).
"""
from pathlib import Path

from ..logging_config import get_logger, log_call
from ..strings import Editing as S

logger = get_logger(__name__)

try:
    from docx import Document as DocxDocument   # noqa: F401 — only to check the lib is present
except ImportError:
    logger.debug("python-docx not installed — Word editing will raise if used.")
    DocxDocument = None  # type: ignore

try:
    import openpyxl
except ImportError:
    logger.debug("openpyxl not installed — Excel editing will raise if used.")
    openpyxl = None  # type: ignore

try:
    import pptx as _pptx_lib   # noqa: F401 — only to check the lib is present
except ImportError:
    logger.debug("python-pptx not installed — PPT editing will raise if used.")
    _pptx_lib = None  # type: ignore

from ..modules.excel import ops as excel_ops
from ..modules.image import ops as image_ops
from ..modules.pptx import ops as pptx_ops
from ..modules.word import ops as word_ops
from ..modules.excel.ops import ExcelOpError
from ..modules.image.ops import ImageOpError
from ..modules.pptx.ops import PptxOpError
from ..modules.word.ops import WordOpError

# Only open formats (OOXML) can be edited (Word/Excel/PPT) — the legacy
# .doc/.xls/.ppt cannot. Images use common raster formats, all directly
# editable (no "open format" constraint like OOXML).
EDITABLE_EXT = {
    ".docx": "word", ".xlsx": "excel", ".pptx": "ppt",
    ".png": "image", ".jpg": "image", ".jpeg": "image",
    ".bmp": "image", ".webp": "image",
}

_MISSING_LIB_TARGET_EXT = {
    "word": ".docx", "excel": ".xlsx", "ppt": ".pptx", "image": ".png/.jpg",
}


class EditError(Exception):
    """File can't be edited — message shown to the user in Vietnamese."""


@log_call
def check_editable(path: str, file_type: str | None) -> None:
    """Raises EditError if the file is outside the editable scope."""
    if file_type not in ("word", "excel", "ppt", "image"):
        raise EditError(S.TYPE_NOT_EDITABLE)

    suffix = Path(path).suffix.lower()
    if suffix not in EDITABLE_EXT:
        raise EditError(S.FORMAT_NOT_EDITABLE.format(
            suffix=suffix, target_ext=_MISSING_LIB_TARGET_EXT[file_type]))
    if file_type == "word" and DocxDocument is None:
        raise EditError(S.MISSING_DOCX_LIB)
    if file_type == "excel" and openpyxl is None:
        raise EditError(S.MISSING_OPENPYXL_LIB)
    if file_type == "ppt" and _pptx_lib is None:
        raise EditError(S.MISSING_PPTX_LIB)


@log_call
def document_outline(path: str, file_type: str | None) -> str:
    """Describe the file's structure to send to the server as context for Claude."""
    try:
        if file_type == "word":
            return word_ops.outline_text(path)
        if file_type == "excel":
            return excel_ops.outline_text(path)
        if file_type == "ppt":
            return pptx_ops.outline_text(path)
        if file_type == "image":
            return image_ops.outline_text(path)
    except EditError:
        raise
    except Exception as exc:  # noqa: BLE001 — corrupted/unreadable file
        raise EditError(S.OUTLINE_READ_FAILED.format(error=exc))
    return ""


@log_call
def apply_edits(path: str, file_type: str | None, edits: list[dict]) -> list[str]:
    """Execute the list of edit commands. Returns a description of each
    command (to report in the chat)."""
    check_editable(path, file_type)
    if not edits:
        return []

    try:
        if file_type == "word":
            return word_ops.apply_edits(path, edits)
        if file_type == "ppt":
            return pptx_ops.apply_edits(path, edits)
        if file_type == "image":
            return image_ops.apply_edits(path, edits)
        return excel_ops.apply_edits(path, edits)
    except (WordOpError, ExcelOpError, PptxOpError, ImageOpError) as exc:
        raise EditError(str(exc))
