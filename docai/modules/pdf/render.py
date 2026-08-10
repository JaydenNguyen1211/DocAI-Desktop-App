import os
import tempfile

from ...imkit import LazyPdfSource

from ...logging_config import get_logger, log_call

logger = get_logger(__name__)


@log_call
def pdf_page_source(path: str, dpi: int = 150) -> LazyPdfSource:
    tmp_dir = tempfile.mkdtemp(prefix="docai_pages_")
    return LazyPdfSource(os.path.abspath(path), tmp_dir, dpi)
