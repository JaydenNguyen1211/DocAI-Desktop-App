"""Xử lý ảnh: (1) chuẩn bị ảnh gửi lên Claude Vision (OCR/trích xuất/tóm tắt),
(2) bộ thao tác sửa ảnh qua chat (xoay/cắt/đổi kích thước/nén).

`prepare_for_vision()` dùng chung cho MỌI lệnh gọi Claude Vision — đây chính là
"tự chỉnh ảnh trước khi OCR" (xoay đúng chiều theo EXIF, tăng tương phản, giới
hạn kích thước): chạy tự động ở tầng hạ tầng, không phải một nút riêng.
"""
import base64
import io
from pathlib import Path

from PIL import Image, ImageOps

_VISION_MAX_EDGE = 2000  # px — dưới ngưỡng high-res 2576px của Claude, vẫn đủ nét để đọc chữ
_VISION_MEDIA_TYPE = {"JPEG": "image/jpeg", "PNG": "image/png"}


class ImageOpError(Exception):
    """Lệnh sửa ảnh không hợp lệ — thông báo tiếng Việt cho người dùng."""


# Các thao tác được phép — cũng là danh sách server mô tả cho Claude.
OPS = ("rotate", "crop", "resize", "compress")


# ── Chuẩn bị ảnh gửi Claude Vision ───────────────────────────────────────────

def prepare_for_vision(path: str) -> tuple[bytes, str]:
    """Trả về (bytes, media_type) đã tự chỉnh — xoay đúng chiều theo EXIF, tăng
    tương phản nhẹ, giới hạn cạnh dài — dùng trước khi gửi ảnh lên Claude."""
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


def b64_for_vision(path: str) -> tuple[str, str]:
    """Như prepare_for_vision() nhưng trả base64 sẵn để đưa vào JSON gửi server."""
    data, media_type = prepare_for_vision(path)
    return base64.b64encode(data).decode("ascii"), media_type


# ── Đọc cấu trúc ───────────────────────────────────────────────────────────

def outline_text(path: str) -> str:
    img = Image.open(path)
    width, height = img.size
    size_kb = Path(path).stat().st_size / 1024
    fmt = img.format or Path(path).suffix.lstrip(".").upper()
    return f"Ảnh {fmt}, kích thước {width}×{height}px, dung lượng {size_kb:.0f}KB."


# ── Thực thi lệnh sửa ──────────────────────────────────────────────────────

def _need(edit: dict, key: str):
    val = edit.get(key)
    if val is None:
        raise ImageOpError(f"Lệnh «{edit.get('op')}» thiếu thông tin: {key}.")
    return val


def apply_edits(path: str, edits: list[dict]) -> list[str]:
    """Áp lần lượt các lệnh sửa lên ảnh, lưu một lần. Trả về mô tả từng lệnh."""
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


def _run_op(img: Image.Image, op: str, edit: dict):
    if op == "rotate":
        angle = float(_need(edit, "angle"))
        img2 = img.rotate(-angle, expand=True)  # âm để khớp chiều kim đồng hồ theo cảm nhận người dùng
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

    # resize — op còn lại trong OPS (compress đã xử lý riêng ở apply_edits)
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
