"""Chuyển đổi định dạng thật — thay `ConvertDialog` cũ (mock `QTimer`, không hề
đọc/ghi file nào cả, xem lịch sử `app/ui/modals.py`). Mỗi hàm nhận đường dẫn nguồn
+ đường dẫn đích tường minh, trả về đường dẫn đích khi thành công, ném
`ConvertError` (thông báo tiếng Việt) khi thất bại.

`ConvertError` kế thừa `ApiError` để `app.thread_worker.CallWorker` tự phát đúng
thông báo (không bọc thêm tiền tố "Lỗi không xác định:") — xem cách
`CallWorker.run()` phân biệt `ApiError` với lỗi khác.

Phần lớn hàm chạy cục bộ, không cần mạng — riêng `pdf_to_excel_file()` cần gọi
AI (Claude đọc PDF trực tiếp qua `/extract_table`) nên cần mạng + quota.
"""
import base64
from pathlib import Path

from PIL import Image

from ...account import api_client
from ..excel.com import ExcelComError, export_workbook_pdf
from ..image import ops as image_ops
from ..pptx.render import ppt_to_pdf
from ..word.render import word_to_pdf
from .creators import create_excel_from_text
from .extractors import pdf_page_texts

_IMG_FORMAT_BY_EXT = {
    ".png": "PNG", ".jpg": "JPEG", ".jpeg": "JPEG",
    ".webp": "WEBP", ".bmp": "BMP",
}


class ConvertError(api_client.ApiError):
    """Không chuyển đổi được — thông báo tiếng Việt cho người dùng."""

    def __init__(self, message: str):
        super().__init__(message, "convert_failed")


# ── Word/Excel/PPT → PDF (COM) ──────────────────────────────────────────────

def word_to_pdf_file(path: str, out_path: str) -> str:
    try:
        return word_to_pdf(path, out_path)
    except Exception as exc:  # noqa: BLE001 — mọi lỗi COM đều quy về 1 thông báo
        raise ConvertError(f"Không chuyển được Word sang PDF: {exc}")


def ppt_to_pdf_file(path: str, out_path: str) -> str:
    try:
        return ppt_to_pdf(path, out_path)
    except Exception as exc:  # noqa: BLE001
        raise ConvertError(f"Không chuyển được PowerPoint sang PDF: {exc}")


def excel_to_pdf_file(path: str, out_path: str) -> str:
    try:
        return export_workbook_pdf(path, out_path)
    except ExcelComError as exc:
        raise ConvertError(str(exc))


# ── PDF → Word/Excel/Ảnh ─────────────────────────────────────────────────────

def pdf_to_word_file(path: str, out_path: str) -> str:
    """Tái tạo nội dung cơ bản: đọc đúng thứ tự chữ, KHÔNG giữ bố cục/style
    gốc. Chỉ hoạt động với PDF có lớp chữ — PDF quét ảnh thuần túy nên chuyển
    qua luồng OCR ảnh (Xử lý ảnh) thay vì chức năng này."""
    pages = pdf_page_texts(path)
    if not any(page_text.strip() for page_text in pages):
        raise ConvertError(
            "PDF này không có lớp chữ (có thể là bản quét ảnh) — không tự "
            "chuyển sang Word được. Hãy dùng OCR ảnh (mở từng trang dạng ảnh) "
            "thay vì chức năng này.")

    from docx import Document as DocxDocument
    doc = DocxDocument()
    for page_index, text in enumerate(pages):
        if page_index > 0:
            doc.add_page_break()
        lines = [line for line in text.split("\n") if line.strip()] or ["(trang trống)"]
        for line in lines:
            doc.add_paragraph(line)
    try:
        doc.save(out_path)
    except PermissionError:
        raise ConvertError(
            f"Không ghi được «{Path(out_path).name}» — file đích đang mở ở chương trình khác.")
    return out_path


def pdf_to_excel_file(path: str, out_path: str) -> str:
    """Nhờ Claude đọc trực tiếp PDF (kể cả bản quét ảnh) rồi trích xuất bảng —
    cần mạng + quota. Dùng lại đúng quy ước "--- Sheet: X ---" + CSV mà
    `create_excel_from_text()` đã parse sẵn cho luồng "tạo tài liệu mới"."""
    data = Path(path).read_bytes()
    attachment = {
        "kind": "pdf", "media_type": "application/pdf",
        "data_b64": base64.b64encode(data).decode("ascii"),
    }
    result = api_client.extract_table(attachment)
    text = (result.get("text") or "").strip()
    if not text:
        raise ConvertError("AI không trích xuất được dữ liệu dạng bảng từ file này.")
    create_excel_from_text(text, out_path)
    return out_path


def pdf_to_images(path: str, out_dir: str) -> list[str]:
    import fitz
    Path(out_dir).mkdir(parents=True, exist_ok=True)
    doc = fitz.open(path)
    out_paths: list[str] = []
    try:
        for page_index, page in enumerate(doc):
            pix = page.get_pixmap(dpi=150)
            page_path = str(Path(out_dir) / f"trang_{page_index + 1:03d}.png")
            pix.save(page_path)
            out_paths.append(page_path)
    finally:
        doc.close()
    if not out_paths:
        raise ConvertError("PDF không có trang nào để xuất.")
    return out_paths


# ── Ảnh → PDF / đổi định dạng ────────────────────────────────────────────────

def image_to_pdf_file(paths: list[str], out_path: str) -> str:
    if not paths:
        raise ConvertError("Không có ảnh nào để chuyển đổi.")
    images = []
    for image_path in paths:
        img = Image.open(image_path)
        if img.mode != "RGB":
            img = img.convert("RGB")
        images.append(img)
    try:
        images[0].save(out_path, "PDF", save_all=True, append_images=images[1:])
    except Exception as exc:  # noqa: BLE001
        raise ConvertError(f"Không tạo được PDF từ ảnh: {exc}")
    return out_path


def image_to_word_file(paths: list[str], out_path: str) -> str:
    """OCR từng ảnh qua Claude Vision (`/extract_text`) rồi dựng thành 1 file
    Word chỉnh sửa được — mỗi ảnh 1 trang riêng (ngắt trang giữa các ảnh), giữ
    đúng cách ngắt đoạn AI đọc được. Cần mạng + quota, tốn 1 lượt gọi/ảnh."""
    if not paths:
        raise ConvertError("Không có ảnh nào để chuyển đổi.")
    from docx import Document as DocxDocument
    doc = DocxDocument()
    got_any = False
    for page_index, path in enumerate(paths):
        if page_index > 0:
            doc.add_page_break()
        data_b64, media_type = image_ops.b64_for_vision(path)
        attachment = {"kind": "image", "media_type": media_type, "data_b64": data_b64}
        result = api_client.extract_text(attachment)
        text = (result.get("text") or "").strip()
        if text and "Khong doc duoc" not in text:
            got_any = True
        for para in (text or "(không đọc được nội dung)").split("\n\n"):
            doc.add_paragraph(para.strip())
    if not got_any:
        raise ConvertError("AI không đọc được nội dung văn bản từ (các) ảnh này.")
    try:
        doc.save(out_path)
    except PermissionError:
        raise ConvertError(
            f"Không ghi được «{Path(out_path).name}» — file đích đang mở ở chương trình khác.")
    return out_path


def image_to_searchable_pdf_file(paths: list[str], out_path: str) -> str:
    """Xuất PDF từ ảnh chụp kèm lớp văn bản OCR ẩn (`render_mode=3` — không vẽ
    ra nhưng vẫn chọn/tìm/copy được) — khác `image_to_pdf_file()` chỉ nhúng
    ảnh thô. Lớp text phủ theo trang, KHÔNG khớp chính xác vị trí từng chữ vì
    Claude Vision trả về văn bản thuần chứ không kèm tọa độ — vẫn đủ để tìm
    kiếm/copy toàn văn. Cần mạng + quota, tốn 1 lượt gọi/ảnh."""
    if not paths:
        raise ConvertError("Không có ảnh nào để chuyển đổi.")
    import fitz
    doc = fitz.open()
    try:
        for path in paths:
            img = Image.open(path)
            width, height = img.size
            page = doc.new_page(width=width, height=height)
            page.insert_image(page.rect, filename=path)

            data_b64, media_type = image_ops.b64_for_vision(path)
            attachment = {"kind": "image", "media_type": media_type, "data_b64": data_b64}
            result = api_client.extract_text(attachment)
            text = (result.get("text") or "").strip()
            if text and "Khong doc duoc" not in text:
                page.insert_textbox(page.rect, text, fontsize=8, render_mode=3)
        doc.save(out_path)
    except Exception as exc:  # noqa: BLE001
        raise ConvertError(f"Không tạo được PDF có thể tìm kiếm: {exc}")
    finally:
        doc.close()
    return out_path


def image_to_excel_file(path: str, out_path: str) -> str:
    """Nhờ Claude đọc ảnh (hóa đơn/hợp đồng chụp) qua Vision rồi trích xuất
    bảng — cần mạng + quota. Cùng /extract_table và cùng quy ước "--- Sheet:
    X ---" + CSV như pdf_to_excel_file()."""
    data_b64, media_type = image_ops.b64_for_vision(path)
    attachment = {"kind": "image", "media_type": media_type, "data_b64": data_b64}
    result = api_client.extract_table(attachment)
    text = (result.get("text") or "").strip()
    if not text:
        raise ConvertError("AI không trích xuất được dữ liệu dạng bảng từ ảnh này.")
    create_excel_from_text(text, out_path)
    return out_path


def image_convert_format(path: str, out_path: str) -> str:
    target_fmt = _IMG_FORMAT_BY_EXT.get(Path(out_path).suffix.lower())
    if target_fmt is None:
        raise ConvertError(f"Định dạng đích chưa hỗ trợ: {Path(out_path).suffix}.")
    try:
        img = Image.open(path)
        if target_fmt == "JPEG" and img.mode != "RGB":
            img = img.convert("RGB")
        img.save(out_path, target_fmt)
    except Exception as exc:  # noqa: BLE001
        raise ConvertError(f"Không chuyển được định dạng ảnh: {exc}")
    return out_path


# ── Điểm vào chung cho ConvertDialog ─────────────────────────────────────────

def convert_file(path: str, source_type: str, target_ext: str, out_path: str,
                  extra_paths: list[str] | None = None, searchable: bool = False) -> str:
    """`source_type`: "word"|"excel"|"ppt"|"pdf"|"image". Trả về đường dẫn kết
    quả — với PDF→Ảnh, `out_path` là THƯ MỤC chứa các trang PNG.

    `extra_paths` (chỉ áp dụng khi source_type="image", target .pdf/.docx):
    ảnh thêm ngoài `path` để gộp thành 1 file nhiều trang theo đúng thứ tự.
    `searchable` (chỉ áp dụng khi target .pdf): True → PDF có lớp text OCR ẩn
    (`image_to_searchable_pdf_file`), False → nhúng ảnh thô như cũ.
    """
    if source_type in ("word", "excel", "ppt") and target_ext == ".pdf":
        fn = {"word": word_to_pdf_file, "excel": excel_to_pdf_file, "ppt": ppt_to_pdf_file}[source_type]
        return fn(path, out_path)
    if source_type == "pdf" and target_ext == ".docx":
        return pdf_to_word_file(path, out_path)
    if source_type == "pdf" and target_ext == ".xlsx":
        return pdf_to_excel_file(path, out_path)
    if source_type == "pdf" and target_ext == ".png":
        pdf_to_images(path, out_path)
        return out_path
    if source_type == "image" and target_ext == ".pdf":
        all_paths = [path] + (extra_paths or [])
        if searchable:
            return image_to_searchable_pdf_file(all_paths, out_path)
        return image_to_pdf_file(all_paths, out_path)
    if source_type == "image" and target_ext == ".docx":
        return image_to_word_file([path] + (extra_paths or []), out_path)
    if source_type == "image" and target_ext == ".xlsx":
        return image_to_excel_file(path, out_path)
    if source_type == "image" and target_ext in _IMG_FORMAT_BY_EXT:
        return image_convert_format(path, out_path)
    raise ConvertError(f"Chưa hỗ trợ chuyển từ {source_type} sang {target_ext}.")
