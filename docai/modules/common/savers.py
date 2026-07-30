import csv
import io
import re

try:
    from docx import Document as DocxDocument
    from docx.shared import Pt
except ImportError:
    DocxDocument = None  # type: ignore
    Pt = None            # type: ignore

try:
    import openpyxl
except ImportError:
    openpyxl = None  # type: ignore


def save_word(original_path: str, edited_text: str, save_path: str):
    doc = DocxDocument()
    try:
        orig = DocxDocument(original_path)
        doc.styles["Normal"].font.name = orig.styles["Normal"].font.name or "Calibri"
        doc.styles["Normal"].font.size = orig.styles["Normal"].font.size or Pt(11)
    except Exception:
        pass
    for line in edited_text.split("\n"):
        doc.add_paragraph(line)
    doc.save(save_path)


def save_excel(edited_text: str, save_path: str):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = None
    for line in edited_text.splitlines():
        match = re.match(r"^---\s*Sheet:\s*(.+?)\s*---\s*$", line)
        if match:
            ws = wb.create_sheet(title=match.group(1)[:31])
            continue
        if ws is None:
            ws = wb.create_sheet(title="Sheet1")
        if line.strip():
            for row in csv.reader(io.StringIO(line)):
                ws.append(row)
    wb.save(save_path)
