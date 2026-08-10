"""Orchestration phi-UI giữa `app.ui.main_window` và tầng `modules`/`pipeline`
— tách riêng để logic gọi mạng + áp lệnh sửa file không nằm lẫn trong 1 QWidget.
"""
import base64
from pathlib import Path
from typing import Optional

from ..account import api_client
from ..modules.image import ops as image_ops
from ..pipeline.editing import EditError, apply_edits, document_outline

from ..logging_config import get_logger, log_call

logger = get_logger(__name__)

_MAX_PDF_ATTACH_BYTES = 30 * 1024 * 1024


@log_call(log_result=False)  # kết quả chứa base64 toàn bộ file — quá lớn để log
def build_attachment(path: str, file_type: str) -> Optional[dict]:
    """Đọc file PDF/ảnh đang mở, chuẩn bị base64 để gửi Claude Vision đọc trực
    tiếp. Trả None nếu không đọc được hoặc không phải loại file hỗ trợ — khi
    đó server vẫn trả lời được, chỉ là không có nội dung file để tham chiếu."""
    try:
        if file_type == "pdf":
            data = Path(path).read_bytes()
            if len(data) > _MAX_PDF_ATTACH_BYTES:
                logger.warning(
                    "PDF attachment too large, skipping: path=%s size=%d bytes (max %d)",
                    path, len(data), _MAX_PDF_ATTACH_BYTES)
                return None
            return {
                "kind": "pdf", "media_type": "application/pdf",
                "data_b64": base64.b64encode(data).decode("ascii"),
            }
        if file_type == "image":
            data_b64, media_type = image_ops.b64_for_vision(path)
            return {"kind": "image", "media_type": media_type, "data_b64": data_b64}
    except Exception:
        logger.exception(
            "Failed to build attachment: path=%s file_type=%s", path, file_type)
        return None
    return None


@log_call
def request_edit(text: str, path: str, ftype: str, history: list) -> dict:
    """Chạy ở luồng nền: gửi cấu trúc tài liệu + yêu cầu lên server, rồi áp
    các lệnh sửa Claude trả về. `edits` rỗng → không phải yêu cầu sửa."""
    try:
        outline = document_outline(path, ftype)
    except EditError as exc:
        raise api_client.ApiError(str(exc), "outline_failed")

    # Ảnh không có "cấu trúc" dạng văn bản như Word/Excel — kèm luôn nội
    # dung ảnh để Claude vẫn trả lời được câu hỏi không phải lệnh sửa
    # (VD "ảnh này chụp gì") thay vì chỉ đọc outline kích thước/dung lượng.
    attachment = build_attachment(path, ftype) if ftype == "image" else None

    data = api_client.edit_file(text, Path(path).name, ftype, outline, history,
                                attachment=attachment)

    edits = data.get("edits") or []
    if not edits:
        return {**data, "notes": []}
    try:
        notes = apply_edits(path, ftype, edits)
    except EditError as exc:
        raise api_client.ApiError(str(exc), "edit_failed")
    return {**data, "notes": notes}
