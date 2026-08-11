"""The set of PowerPoint (.pptx) edit operations for the edit-via-chat flow.

Mirrors `word_ops.py`: `outline_text()` describes slides/shapes so Claude can
target the right position; `apply_edits()` executes the list of edit
commands Claude returns.

python-pptx has no built-in API to move/delete slides — operate directly on
`prs.slides._sldIdLst` (the slide list in the XML), the common approach with
this library.
"""
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Cm, Pt

from ...logging_config import get_logger, log_call

logger = get_logger(__name__)


class PptxOpError(Exception):
    """Invalid edit command — message shown to the user in Vietnamese."""


# Allowed operations — also the list the server describes to Claude.
OPS = (
    "set_text",       # reset a shape's entire text content
    "format_text",    # font/size/color/bold/italic/underline/alignment
    "set_list",       # set/unset a bullet or numbered list
    "add_slide",      # add a new slide
    "delete_slide",   # delete a slide
    "move_slide",     # change a slide's position
    "insert_image",   # insert an image (from a real local file) into a slide
)

_LAYOUT_MAP = {"title": 0, "title_content": 1, "two_content": 3, "blank": 6}
_ALIGN = {
    "left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
    "right": PP_ALIGN.RIGHT, "justify": PP_ALIGN.JUSTIFY,
}


@log_call
def _need(edit: dict, key: str):
    val = edit.get(key)
    if val is None or (isinstance(val, str) and not val.strip()):
        raise PptxOpError(f"Lệnh «{edit.get('op')}» thiếu thông tin: {key}.")
    return val


@log_call
def _short(text, limit: int = 80) -> str:
    text = " ".join(str(text).split())
    return text if len(text) <= limit else text[:limit] + "…"


@log_call
def _clean_hex(color) -> str:
    hex_color = str(color).strip().lstrip("#").upper()
    if len(hex_color) != 6 or any(ch not in "0123456789ABCDEF" for ch in hex_color):
        raise PptxOpError(f"Mã màu không hợp lệ: {color}.")
    return hex_color


# ── Read structure ────────────────────────────────────────────────────────

@log_call
def outline_text(path: str) -> str:
    """A compact text-form outline for the prompt."""
    prs = Presentation(path)
    lines = [f"Bài trình bày có {len(prs.slides)} slide."]
    for si, slide in enumerate(prs.slides):
        lines.append(f"\n[Slide {si}] layout «{slide.slide_layout.name}»")
        has_content = False
        for shi, shape in enumerate(slide.shapes):
            if shape.has_text_frame and shape.text_frame.text.strip():
                has_content = True
                is_title = shape.is_placeholder and shape.placeholder_format.idx == 0
                kind = "tiêu đề" if is_title else "nội dung"
                lines.append(f"  [{shi}] ({kind}) {_short(shape.text_frame.text)}")
            elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                has_content = True
                lines.append(f"  [{shi}] (ảnh)")
        if not has_content:
            lines.append("  (trống)")
    return "\n".join(lines)


# ── Helpers ───────────────────────────────────────────────────────────────────

@log_call
def _slide_at(prs, edit: dict, key: str = "slide"):
    idx = int(_need(edit, key))
    slide_count = len(prs.slides)
    if not 0 <= idx < slide_count:
        raise PptxOpError(f"Bài trình bày chỉ có {slide_count} slide (0–{slide_count - 1}), không có slide {idx}.")
    return prs.slides[idx]


@log_call
def _shape_at(slide, edit: dict):
    idx = int(_need(edit, "shape_index"))
    shapes = list(slide.shapes)
    if not 0 <= idx < len(shapes):
        raise PptxOpError(
            f"Slide chỉ có {len(shapes)} đối tượng (0–{len(shapes) - 1}), không có shape {idx}.")
    shape = shapes[idx]
    if not shape.has_text_frame:
        raise PptxOpError(f"Shape {idx} không có nội dung chữ để sửa.")
    return shape


@log_call
def _fill_text_frame(tf, text: str):
    lines = text.split("\n")
    tf.text = lines[0]
    for line in lines[1:]:
        tf.add_paragraph().text = line


@log_call
def _move_slide_element(prs, from_pos: int, to_pos: int):
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    el = slides[from_pos]
    xml_slides.remove(el)
    xml_slides.insert(to_pos, el)


# ── Execute edit commands ────────────────────────────────────────────────────

@log_call
def apply_edits(path: str, edits: list[dict]) -> list[str]:
    """Apply the edit commands to the file in order, save once. Returns a
    description of each command."""
    if not edits:
        return []

    prs = Presentation(path)
    notes: list[str] = []
    for edit in edits:
        op = (edit.get("op") or "").strip()
        if op not in OPS:
            raise PptxOpError(f"Thao tác chưa hỗ trợ: «{op}».")
        notes.append(_run_op(prs, op, edit))

    try:
        prs.save(path)
    except PermissionError:
        raise PptxOpError(
            f"Không ghi được «{Path(path).name}» — file đang mở trong PowerPoint, "
            "hãy đóng lại rồi thử lại.")
    return notes


@log_call
def _run_op(prs, op: str, edit: dict) -> str:
    if op == "add_slide":
        return _add_slide(prs, edit)
    if op == "delete_slide":
        return _delete_slide(prs, edit)
    if op == "move_slide":
        return _move_slide(prs, edit)

    slide = _slide_at(prs, edit)

    if op == "insert_image":
        return _insert_image(slide, edit)

    shape = _shape_at(slide, edit)
    if op == "set_text":
        return _set_text(shape, edit)
    if op == "format_text":
        return _format_text(shape, edit)

    # set_list — the last remaining op in OPS
    return _set_list(shape, edit)


@log_call
def _set_text(shape, edit: dict) -> str:
    text = _need(edit, "text")
    tf = shape.text_frame
    for paragraph in list(tf.paragraphs[1:]):
        paragraph._p.getparent().remove(paragraph._p)
    _fill_text_frame(tf, text)
    return f"đặt nội dung shape {edit['shape_index']} ở slide {edit['slide']} thành «{_short(text)}»"


@log_call
def _format_text(shape, edit: dict) -> str:
    tf = shape.text_frame
    target = edit.get("text") or None
    paragraphs = list(tf.paragraphs)

    if target:
        paragraphs = [paragraph for paragraph in paragraphs if target in paragraph.text]
        if not paragraphs:
            raise PptxOpError(
                f"Không tìm thấy «{_short(target)}» trong shape {edit['shape_index']}.")

    changes = []
    for key, label in (("bold", "in đậm"), ("italic", "in nghiêng"), ("underline", "gạch chân")):
        if edit.get(key) is not None:
            changes.append(label if edit[key] else f"bỏ {label}")
    if edit.get("font_size"):
        changes.append(f"cỡ chữ {float(edit['font_size']):g}pt")
    if edit.get("font_name"):
        changes.append(f"font «{edit['font_name']}»")
    if edit.get("font_color"):
        changes.append(f"màu chữ #{str(edit['font_color']).lstrip('#').upper()}")
    if edit.get("align"):
        changes.append(f"căn {edit['align']}")
    if not changes:
        raise PptxOpError("Lệnh «format_text» không nêu định dạng nào.")

    for paragraph in paragraphs:
        runs = list(paragraph.runs) or [paragraph.add_run()]
        for key in ("bold", "italic", "underline"):
            if edit.get(key) is not None:
                val = bool(edit[key])
                for run in runs:
                    setattr(run, key, val)
        if edit.get("font_size"):
            size = float(edit["font_size"])
            for run in runs:
                run.font.size = Pt(size)
        if edit.get("font_name"):
            name = str(edit["font_name"])
            for run in runs:
                run.font.name = name
        if edit.get("font_color"):
            hex_color = _clean_hex(edit["font_color"])
            for run in runs:
                run.font.color.rgb = RGBColor.from_string(hex_color)
        if edit.get("align"):
            align = str(edit["align"]).lower()
            if align not in _ALIGN:
                raise PptxOpError(f"Kiểu căn lề chưa hỗ trợ: {align}.")
            paragraph.alignment = _ALIGN[align]

    scope = (f"«{_short(target)}» trong shape {edit['shape_index']}" if target
            else f"shape {edit['shape_index']}")
    return f"định dạng {scope} ở slide {edit['slide']}: {', '.join(changes)}"


@log_call
def _set_list(shape, edit: dict) -> str:
    list_type = str(_need(edit, "list_type")).lower()
    if list_type not in ("bullet", "number", "none"):
        raise PptxOpError('list_type phải là "bullet", "number" hoặc "none".')

    for paragraph in shape.text_frame.paragraphs:
        pPr = paragraph._p.get_or_add_pPr()
        for tag in ("a:buChar", "a:buAutoNum", "a:buNone"):
            el = pPr.find(qn(tag))
            if el is not None:
                pPr.remove(el)
        if list_type == "bullet":
            etree.SubElement(pPr, qn("a:buChar"), {"char": "•"})
        elif list_type == "number":
            etree.SubElement(pPr, qn("a:buAutoNum"), {"type": "arabicPeriod"})
        else:
            etree.SubElement(pPr, qn("a:buNone"))

    label = {"bullet": "gạch đầu dòng", "number": "đánh số", "none": "bỏ danh sách"}[list_type]
    return f"đặt {label} cho shape {edit['shape_index']} ở slide {edit['slide']}"


@log_call
def _add_slide(prs, edit: dict) -> str:
    layout_key = str(_need(edit, "layout")).lower()
    if layout_key not in _LAYOUT_MAP:
        raise PptxOpError('layout phải là "title", "title_content", "two_content" hoặc "blank".')
    layout_idx = _LAYOUT_MAP[layout_key]
    if layout_idx >= len(prs.slide_layouts):
        raise PptxOpError(f"Mẫu trình bày không có layout tương ứng «{layout_key}».")

    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])

    title = edit.get("title")
    if title and slide.shapes.title is not None:
        slide.shapes.title.text = str(title)

    body_text = edit.get("body")
    if body_text:
        body_ph = next(
            (ph for ph in slide.placeholders if ph.placeholder_format.idx != 0), None)
        if body_ph is not None:
            _fill_text_frame(body_ph.text_frame, str(body_text))

    new_pos = len(prs.slides) - 1
    where = f"cuối bài trình bày (slide {new_pos})"
    index = edit.get("index")
    if index is not None:
        index = int(index)
        if not 0 <= index <= new_pos:
            raise PptxOpError(
                f"Vị trí chèn không hợp lệ: {index} (có {new_pos} slide trước khi thêm).")
        _move_slide_element(prs, new_pos, index)
        where = f"vị trí {index}"

    return f"thêm slide mới (layout «{layout_key}») vào {where}"


@log_call
def _delete_slide(prs, edit: dict) -> str:
    idx = int(_need(edit, "index"))
    slide_count = len(prs.slides)
    if not 0 <= idx < slide_count:
        raise PptxOpError(f"Bài trình bày chỉ có {slide_count} slide, không có slide {idx}.")
    if slide_count == 1:
        raise PptxOpError("Không thể xóa slide cuối cùng của bài trình bày.")

    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    rId = slides[idx].rId
    prs.part.drop_rel(rId)
    xml_slides.remove(slides[idx])
    return f"xóa slide {idx}"


@log_call
def _move_slide(prs, edit: dict) -> str:
    slide_count = len(prs.slides)
    from_idx = int(_need(edit, "from_index"))
    to_idx = int(_need(edit, "to_index"))
    if not 0 <= from_idx < slide_count:
        raise PptxOpError(f"Bài trình bày chỉ có {slide_count} slide, không có slide {from_idx}.")
    if not 0 <= to_idx < slide_count:
        raise PptxOpError(f"Bài trình bày chỉ có {slide_count} slide, không có vị trí {to_idx}.")
    _move_slide_element(prs, from_idx, to_idx)
    return f"chuyển slide {from_idx} tới vị trí {to_idx}"


@log_call
def _insert_image(slide, edit: dict) -> str:
    image_path = _need(edit, "image_path")
    if not Path(image_path).is_file():
        raise PptxOpError(f"Không tìm thấy file ảnh: {image_path}.")

    left = Cm(float(edit["left"])) if edit.get("left") is not None else Cm(2)
    top = Cm(float(edit["top"])) if edit.get("top") is not None else Cm(2)
    kwargs = {}
    if edit.get("width") is not None:
        kwargs["width"] = Cm(float(edit["width"]))

    try:
        slide.shapes.add_picture(str(image_path), left, top, **kwargs)
    except Exception as exc:  # noqa: BLE001 — corrupted/unreadable image file
        raise PptxOpError(f"Không chèn được ảnh: {exc}")
    return f"chèn ảnh «{Path(image_path).name}» vào slide"
