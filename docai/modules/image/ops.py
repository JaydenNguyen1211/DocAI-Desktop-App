"""Image processing: (1) prepare images sent to Claude Vision (OCR/extract/
summarize), (2) the set of image edit operations via chat (rotate/crop/
resize/compress).

`prepare_for_vision()` is shared by EVERY call to Claude Vision — this is the
"auto-adjust the image before OCR" step (rotate correctly per EXIF, boost
contrast, cap size): runs automatically at the infrastructure layer, not a
separate button.
"""
import base64
import io
from pathlib import Path

from PIL import Image, ImageOps

from ...logging_config import get_logger, log_call

logger = get_logger(__name__)

_VISION_MAX_EDGE = 2000  # px — below Claude's 2576px high-res threshold, still sharp enough to read text
_VISION_MEDIA_TYPE = {"JPEG": "image/jpeg", "PNG": "image/png"}


class ImageOpError(Exception):
    """Invalid image edit command — message shown to the user in Vietnamese."""


# Allowed operations — also the list the server describes to Claude.
OPS = ("rotate", "crop", "resize", "compress")


# ── Prepare images sent to Claude Vision ─────────────────────────────────────

@log_call
def prepare_for_vision(path: str) -> tuple[bytes, str]:
    """Returns (bytes, media_type) already auto-adjusted — rotated correctly
    per EXIF, slight contrast boost, long edge capped — used before sending
    the image to Claude."""
    img = Image.open(path)
    img = ImageOps.exif_transpose(img)
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    img = ImageOps.autocontrast(img, cutoff=1)

    width, height = img.size
    if max(width, height) > _VISION_MAX_EDGE:
        scale = _VISION_MAX_EDGE / max(width, height)
        img = img.resize((max(1, round(width * scale)), max(1, round(height * scale))), Image.LANCZOS)

    buf = io.BytesIO()
    fmt = "PNG" if img.mode == "L" else "JPEG"
    if fmt == "JPEG":
        img.save(buf, "JPEG", quality=88)
    else:
        img.save(buf, "PNG")
    return buf.getvalue(), _VISION_MEDIA_TYPE[fmt]


@log_call
def b64_for_vision(path: str) -> tuple[str, str]:
    """Like prepare_for_vision() but returns base64 ready to drop into the JSON sent to the server."""
    data, media_type = prepare_for_vision(path)
    return base64.b64encode(data).decode("ascii"), media_type


# ── Read structure ────────────────────────────────────────────────────────

@log_call
def outline_text(path: str) -> str:
    img = Image.open(path)
    width, height = img.size
    size_kb = Path(path).stat().st_size / 1024
    fmt = img.format or Path(path).suffix.lstrip(".").upper()
    return f"Ảnh {fmt}, kích thước {width}×{height}px, dung lượng {size_kb:.0f}KB."


# ── Execute edit commands ────────────────────────────────────────────────────

@log_call
def _need(edit: dict, key: str):
    val = edit.get(key)
    if val is None:
        raise ImageOpError(f"Lệnh «{edit.get('op')}» thiếu thông tin: {key}.")
    return val


@log_call
def apply_edits(path: str, edits: list[dict]) -> list[str]:
    """Apply the edit commands to the image in order, save once. Returns a
    description of each command."""
    if not edits:
        return []

    img = Image.open(path)
    img = ImageOps.exif_transpose(img)
    quality: int | None = None
    notes: list[str] = []

    for edit in edits:
        op = (edit.get("op") or "").strip()
        if op not in OPS:
            raise ImageOpError(f"Thao tác chưa hỗ trợ: «{op}».")
        if op == "compress":
            quality = int(edit.get("quality") or 70)
            if not 1 <= quality <= 95:
                raise ImageOpError("quality phải trong khoảng 1–95.")
            notes.append(f"nén ảnh (chất lượng {quality})")
            continue
        img, note = _run_op(img, op, edit)
        notes.append(note)

    ext = Path(path).suffix.lower()
    save_kwargs: dict = {}
    if ext in (".jpg", ".jpeg"):
        if img.mode != "RGB":
            img = img.convert("RGB")
        save_kwargs["quality"] = quality or 90
        save_kwargs["optimize"] = True
    elif ext == ".webp":
        save_kwargs["quality"] = quality if quality is not None else 80
    elif ext == ".png":
        save_kwargs["optimize"] = True

    try:
        img.save(path, **save_kwargs)
    except PermissionError:
        raise ImageOpError(
            f"Không ghi được «{Path(path).name}» — file đang mở ở chương trình "
            "khác, hãy đóng lại rồi thử lại.")
    return notes


@log_call
def _run_op(img: Image.Image, op: str, edit: dict):
    if op == "rotate":
        angle = float(_need(edit, "angle"))
        img2 = img.rotate(-angle, expand=True)  # negated to match users' intuition of clockwise rotation
        return img2, f"xoay ảnh {angle:g}°"

    if op == "crop":
        left = int(_need(edit, "left"))
        top = int(_need(edit, "top"))
        right = int(_need(edit, "right"))
        bottom = int(_need(edit, "bottom"))
        width, height = img.size
        if not (0 <= left < right <= width and 0 <= top < bottom <= height):
            raise ImageOpError(
                f"Vùng cắt ({left},{top},{right},{bottom}) vượt quá kích thước ảnh ({width}×{height}).")
        img2 = img.crop((left, top, right, bottom))
        return img2, f"cắt ảnh còn vùng ({left},{top})–({right},{bottom})"

    # resize — the remaining op in OPS (compress is handled separately in apply_edits)
    width = edit.get("width")
    height = edit.get("height")
    if not width and not height:
        raise ImageOpError("«resize» cần width và/hoặc height.")
    orig_width, orig_height = img.size
    if width and not height:
        height = round(orig_height * (int(width) / orig_width))
    elif height and not width:
        width = round(orig_width * (int(height) / orig_height))
    img2 = img.resize((int(width), int(height)), Image.LANCZOS)
    return img2, f"đổi kích thước ảnh thành {int(width)}×{int(height)}px"
