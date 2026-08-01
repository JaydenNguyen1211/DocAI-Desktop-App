"""Tiện ích ảnh/trang dùng chung cho các bộ renderer theo định dạng
(word/excel/pptx/image/pdf) — nguồn trang lazy/eager + chuyển đổi PIL↔Qt."""
import os
import threading

from PySide6.QtGui import QPixmap, QImage

try:
    import fitz as _fitz                # noqa: F401
    from PIL import Image as _PILImg    # noqa: F401
    HAS_PAGEPREVIEW = True
except ImportError:
    HAS_PAGEPREVIEW = False


class LazyPdfSource:
    """Nguồn trang PDF render theo yêu cầu.

    Dùng cho PDF gốc, và cho Word/PPT sau khi đã quy về 1 PDF trung gian
    qua COM. Mở file 1 lần lấy ngay tổng số trang + kích thước trang (đều
    là đọc metadata, không raster) — rẻ dù tài liệu vài nghìn trang. Chỉ
    khi `png_path(i)` được gọi thì trang đó mới thực sự được raster ra PNG
    (và cache lại), để widget xem trước chỉ tải trang nào người dùng đang
    thực sự cuộn tới thay vì raster hết cả cuốn sách.
    """

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

    def page_size(self, index: int) -> tuple[int, int]:
        zoom = self._dpi / 72.0
        with self._lock:
            rect = self._doc[index].rect
        return max(1, int(rect.width * zoom)), max(1, int(rect.height * zoom))

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
                pass
        shutil.rmtree(self._tmp_dir, ignore_errors=True)


class EagerPageSource:
    """Nguồn trang đã raster sẵn hết — dùng cho Excel/ảnh, vốn rất ít trang
    (số sheet, hoặc 1) nên không cần tải lười."""

    def __init__(self, png_paths: list[str], tmp_dir: str):
        self._paths = png_paths
        self.page_count = len(png_paths)
        self._tmp_dir = tmp_dir
        self._closed = False

    def page_size(self, index: int) -> tuple[int, int]:
        from PIL import Image
        with Image.open(self._paths[index]) as img:
            return img.size

    def png_path(self, index: int) -> str:
        return self._paths[index]

    def close(self):
        import shutil
        if self._closed:
            return
        self._closed = True
        shutil.rmtree(self._tmp_dir, ignore_errors=True)


def pil_to_qimage(pil_img) -> QImage:
    """An toàn gọi từ luồng nền (khác QPixmap, QImage không cần luồng GUI).
    `.copy()` để QImage sở hữu buffer riêng, tách khỏi bytes Python cục bộ
    trước khi vượt qua ranh giới luồng."""
    pil_img = pil_img.convert("RGB")
    data = pil_img.tobytes("raw", "RGB")
    qimg = QImage(data, pil_img.width, pil_img.height,
                  pil_img.width * 3, QImage.Format.Format_RGB888)
    return qimg.copy()


def pil_to_qpixmap(pil_img) -> QPixmap:
    return QPixmap.fromImage(pil_to_qimage(pil_img))
