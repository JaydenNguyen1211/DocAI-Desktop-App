import os
import tempfile

from ...imkit import LazyPdfSource


def ppt_to_pdf(path: str, out_path: str | None = None) -> str:
    """Xuất .pptx ra PDF qua PowerPoint COM. `out_path` bỏ trống → file tạm
    dùng riêng cho bản xem trước (`ppt_page_source`); truyền vào khi đây là
    thao tác CHUYỂN ĐỔI thật của người dùng (xem `modules/common/converters.py`)."""
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
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        pres = ppt.Presentations.Open(abs_path, ReadOnly=True, WithWindow=False)
        pres.SaveAs(out_path, 32)   # ppSaveAsPDF
        pres.Close()
        ppt.Quit()
    finally:
        pythoncom.CoUninitialize()
    return out_path


def ppt_page_source(path: str, dpi: int = 150) -> LazyPdfSource:
    pdf_path = ppt_to_pdf(path)
    tmp_dir = tempfile.mkdtemp(prefix="docai_pages_")
    return LazyPdfSource(pdf_path, tmp_dir, dpi, owned_files=[pdf_path])
