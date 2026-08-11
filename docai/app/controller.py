"""Non-UI orchestration between `app.ui.main_window` and the `modules`/
`pipeline` layer — kept separate so network calls + applying file edits
aren't mixed into a QWidget.
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


@log_call(log_result=False)  # result contains the whole file as base64 — too large to log
def build_attachment(path: str, file_type: str) -> Optional[dict]:
    """Read the currently open PDF/image file and base64-encode it so Claude
    Vision can read it directly. Returns None if it can't be read or isn't a
    supported file type — the server can still reply, it just won't have any
    file content to reference."""
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
    """Runs on a background thread: sends the document structure + request
    to the server, then applies the edit commands Claude returns. An empty
    `edits` means it wasn't an edit request."""
    try:
        outline = document_outline(path, ftype)
    except EditError as exc:
        raise api_client.ApiError(str(exc), "outline_failed")

    # Images don't have a text "structure" like Word/Excel — attach the
    # image content itself so Claude can still answer questions that aren't
    # edit commands (e.g. "what is this a photo of") instead of only seeing
    # a size/dimensions outline.
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
