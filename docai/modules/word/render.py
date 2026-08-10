import os
import sys
import tempfile

from ...imkit import LazyPdfSource

from ...logging_config import get_logger, log_call

logger = get_logger(__name__)


@log_call
def word_to_pdf(path: str, out_path: str | None = None) -> str:
    """Xuất .docx ra PDF. Windows: Word COM. macOS: Microsoft Office → LibreOffice. Linux: LibreOffice."""
    abs_path = os.path.abspath(path)
    if out_path is None:
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tf:
            out_path = tf.name
    else:
        out_path = os.path.abspath(out_path)

    if sys.platform == "win32":
        _word_to_pdf_com(abs_path, out_path)
    elif sys.platform == "darwin":
        _word_to_pdf_mac(abs_path, out_path)
    else:
        from ..common._soffice import convert_to_pdf
        convert_to_pdf(abs_path, out_path)
    return out_path


@log_call
def _word_to_pdf_mac(abs_path: str, out_path: str) -> None:
    from ..common._msoffice_mac import is_available, word_to_pdf as ms_word_to_pdf
    from ..common._soffice import convert_to_pdf
    if is_available("word"):
        ms_word_to_pdf(abs_path, out_path)
    else:
        convert_to_pdf(abs_path, out_path)


@log_call
def _word_to_pdf_com(abs_path: str, out_path: str) -> None:
    import win32com.client
    import pythoncom
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


@log_call
def word_page_source(path: str, dpi: int = 150) -> LazyPdfSource:
    pdf_path = word_to_pdf(path)
    tmp_dir = tempfile.mkdtemp(prefix="docai_pages_")
    return LazyPdfSource(pdf_path, tmp_dir, dpi, owned_files=[pdf_path])
