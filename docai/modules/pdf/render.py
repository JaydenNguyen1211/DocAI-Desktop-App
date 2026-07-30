import os
import tempfile

from ...imkit import LazyPdfSource


def pdf_page_source(path: str, dpi: int = 150) -> LazyPdfSource:
    tmp_dir = tempfile.mkdtemp(prefix="docai_pages_")
    return LazyPdfSource(os.path.abspath(path), tmp_dir, dpi)
