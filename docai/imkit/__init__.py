"""Shared image/page utilities for the per-format renderers
(word/excel/pptx/image/pdf) — lazy/eager page sources + PIL↔Qt conversion."""
import os
import threading

from PySide6.QtGui import QPixmap, QImage

from ..logging_config import get_logger, log_call

logger = get_logger(__name__)

try:
    import fitz as _fitz                # noqa: F401
    from PIL import Image as _PILImg    # noqa: F401
    HAS_PAGEPREVIEW = True
except ImportError:
    logger.debug("pymupdf/Pillow not installed — page preview unavailable.")
    HAS_PAGEPREVIEW = False


class LazyPdfSource:
    """PDF page source that renders on demand.

    Used for native PDFs, and for Word/PPT after they've been converted to
    an intermediate PDF via COM. The file is opened once to get the page
    count + page sizes right away (both are metadata reads, no rasterizing)
    — cheap even for a document with thousands of pages. A page is only
    actually rasterized to PNG (and cached) when `png_path(i)` is called,
    so the preview widget only loads the pages the user is actually
    scrolling to instead of rasterizing the whole book.
    """

    @log_call
    def __init__(self, pdf_path: str, tmp_dir: str, dpi: int = 150,
                 owned_files: list[str] | None = None):
        import fitz
        self._doc = fitz.open(pdf_path)
        self.page_count = self._doc.page_count
        self._dpi = dpi
        self._tmp_dir = tmp_dir
        self._lock = threading.Lock()
        self._cache: dict[int, str] = {}
        self._owned_files = owned_files or []
        self._closed = False

    @log_call
    def page_size(self, index: int) -> tuple[int, int]:
        zoom = self._dpi / 72.0
        with self._lock:
            rect = self._doc[index].rect
        return max(1, int(rect.width * zoom)), max(1, int(rect.height * zoom))

    @log_call
    def png_path(self, index: int) -> str:
        cached = self._cache.get(index)
        if cached:
            return cached
        import fitz
        zoom = self._dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        with self._lock:
            cached = self._cache.get(index)
            if cached:
                return cached
            page = self._doc[index]
            pix = page.get_pixmap(matrix=mat, alpha=False)
            out = os.path.join(self._tmp_dir, f"page_{index:04d}.png")
            pix.save(out)
            self._cache[index] = out
        return out

    @log_call
    def close(self):
        import shutil
        if self._closed:
            return
        self._closed = True
        with self._lock:
            self._doc.close()
        for owned_file in self._owned_files:
            try:
                os.unlink(owned_file)
            except Exception:
                logger.debug("Could not delete temp owned file: %s", owned_file,
                            exc_info=True)
        shutil.rmtree(self._tmp_dir, ignore_errors=True)


class EagerPageSource:
    """Page source that's already fully rasterized — used for Excel/images,
    which have very few pages (sheet count, or 1) so lazy loading isn't needed."""

    @log_call
    def __init__(self, png_paths: list[str], tmp_dir: str):
        self._paths = png_paths
        self.page_count = len(png_paths)
        self._tmp_dir = tmp_dir
        self._closed = False

    @log_call
    def page_size(self, index: int) -> tuple[int, int]:
        from PIL import Image
        with Image.open(self._paths[index]) as img:
            return img.size

    @log_call
    def png_path(self, index: int) -> str:
        return self._paths[index]

    @log_call
    def close(self):
        import shutil
        if self._closed:
            return
        self._closed = True
        shutil.rmtree(self._tmp_dir, ignore_errors=True)


@log_call
def pil_to_qimage(pil_img) -> QImage:
    """Safe to call from a background thread (unlike QPixmap, QImage doesn't
    need the GUI thread). `.copy()` so the QImage owns its own buffer,
    detached from the local Python bytes, before crossing the thread
    boundary."""
    pil_img = pil_img.convert("RGB")
    data = pil_img.tobytes("raw", "RGB")
    qimg = QImage(data, pil_img.width, pil_img.height,
                  pil_img.width * 3, QImage.Format.Format_RGB888)
    return qimg.copy()


@log_call
def pil_to_qpixmap(pil_img) -> QPixmap:
    return QPixmap.fromImage(pil_to_qimage(pil_img))
