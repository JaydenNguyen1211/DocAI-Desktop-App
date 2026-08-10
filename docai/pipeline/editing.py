"""Áp lệnh sửa của AI lên file Word/Excel/PowerPoint/ảnh ngay tại chỗ (ghi đè
file gốc).

Server (Claude) quyết định danh sách lệnh sửa; module này chỉ điều phối tới
đúng engine trong `modules/` theo loại file.
  · Word    — bộ thao tác đầy đủ trong `modules.word.ops` (thêm/xóa đoạn, xóa
              trang, chèn vào vị trí cụ thể, định dạng, tìm & thay…).
  · Excel   — bộ thao tác đầy đủ trong `modules.excel.ops` (ô/vùng, định
              dạng, bảng, sắp xếp, lọc, biểu đồ, sheet, freeze pane…).
  · PPT     — bộ thao tác trong `modules.pptx.ops` (chữ/định dạng/danh sách,
              thêm-xóa-dời slide, chèn ảnh).
  · Ảnh     — bộ thao tác trong `modules.image.ops` (xoay/cắt/đổi kích
              thước/nén).
"""
from pathlib import Path

from ..logging_config import get_logger, log_call
from ..strings import Editing as S

logger = get_logger(__name__)

try:
    from docx import Document as DocxDocument   # noqa: F401 — chỉ để kiểm tra có lib
except ImportError:
    logger.debug("python-docx not installed — Word editing will raise if used.")
    DocxDocument = None  # type: ignore

try:
    import openpyxl
except ImportError:
    logger.debug("openpyxl not installed — Excel editing will raise if used.")
    openpyxl = None  # type: ignore

try:
    import pptx as _pptx_lib   # noqa: F401 — chỉ để kiểm tra có lib
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

# Chỉ định dạng mở (OOXML) mới sửa được (Word/Excel/PPT) — .doc/.xls/.ppt cũ
# thì không. Ảnh dùng định dạng raster phổ biến, sửa trực tiếp được cả (không
# có ràng buộc "định dạng mở" như OOXML).
EDITABLE_EXT = {
    ".docx": "word", ".xlsx": "excel", ".pptx": "ppt",
    ".png": "image", ".jpg": "image", ".jpeg": "image",
    ".bmp": "image", ".webp": "image",
}

_MISSING_LIB_TARGET_EXT = {
    "word": ".docx", "excel": ".xlsx", "ppt": ".pptx", "image": ".png/.jpg",
}


class EditError(Exception):
    """Không sửa được file — thông báo tiếng Việt cho người dùng."""


@log_call
def check_editable(path: str, file_type: str | None) -> None:
    """Ném EditError nếu file không nằm trong phạm vi sửa được."""
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
    """Mô tả cấu trúc file để gửi lên server làm ngữ cảnh cho Claude."""
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
    except Exception as exc:  # noqa: BLE001 — file hỏng/không đọc được
        raise EditError(S.OUTLINE_READ_FAILED.format(error=exc))
    return ""


@log_call
def apply_edits(path: str, file_type: str | None, edits: list[dict]) -> list[str]:
    """Thực thi danh sách lệnh sửa. Trả về mô tả từng lệnh (để báo trong chat)."""
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
