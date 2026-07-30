import os
import tempfile

from ...imkit import LazyPdfSource


def word_to_pdf(path: str, out_path: str | None = None) -> str:
    """Xuất .docx ra PDF qua Word COM. `out_path` bỏ trống → file tạm dùng
    riêng cho bản xem trước (`word_page_source`); truyền vào khi đây là thao
    tác CHUYỂN ĐỔI thật của người dùng (xem `modules/common/converters.py`)."""
    import win32com.client
    import pythoncom

    abs_path = os.path.abspath(path)
    if out_path is None:
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tf:
            out_path = tf.name
    else:
        out_path = os.path.abspath(out_path)
    pythoncom.CoInitialize()
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(abs_path)
        doc.SaveAs2(out_path, FileFormat=17)
        doc.Close(False)
        word.Quit()
    finally:
        pythoncom.CoUninitialize()
    return out_path


def word_page_source(path: str, dpi: int = 150) -> LazyPdfSource:
    pdf_path = word_to_pdf(path)
    tmp_dir = tempfile.mkdtemp(prefix="docai_pages_")
    return LazyPdfSource(pdf_path, tmp_dir, dpi, owned_files=[pdf_path])
