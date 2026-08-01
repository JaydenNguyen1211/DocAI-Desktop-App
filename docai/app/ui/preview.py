"""Panel xem trước tài liệu — render trang qua Word COM / LibreOffice / fitz / PIL."""
import bisect
import os
import subprocess
import sys
import threading
from collections import OrderedDict
from pathlib import Path

from PySide6.QtCore import Qt, QSize, QTimer, Signal
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import (
    QFrame, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QLineEdit,
    QScrollArea, QWidget,
)

from ...imkit import pil_to_qimage, HAS_PAGEPREVIEW
from ...modules.common.extractors import pdf_page_texts, HAS_PDF_TEXT
from ...modules.excel.render import excel_page_source
from ...modules.image.render import image_page_source
from ...modules.pdf.render import pdf_page_source
from ...modules.pptx.render import ppt_page_source
from ...modules.word.render import word_page_source
from ..constants import FILE_BADGE, OPEN_WITH_LABEL

_SOURCE_FACTORIES = {
    "word": word_page_source,
    "excel": excel_page_source,
    "pdf": pdf_page_source,
    "image": image_page_source,
    "ppt": ppt_page_source,
}


def _open_path(path: str) -> None:
    if sys.platform == "win32":
        os.startfile(path)  # noqa: S606
    elif sys.platform == "darwin":
        subprocess.run(["open", path], check=False)
    else:
        subprocess.run(["xdg-open", path], check=False)


class PagePreviewWidget(QScrollArea):
    """Danh sách trang xem trước — ảo hoá theo viewport.

    Với sách vài trăm–nghìn trang, việc RASTER từng trang và TẠO WIDGET cho
    từng trang đều là việc tốn thời gian tỉ lệ với số trang — làm hết ngay
    lúc mở file (dù ở luồng nền hay luồng chính) khiến người dùng phải chờ
    rất lâu mới thấy gì, hoặc làm đơ UI nếu chạy tuần tự trên luồng chính.

    Cách này chỉ: mở tài liệu lấy số trang + kích thước trang đầu (đọc
    metadata, không raster — rẻ dù tài liệu vài nghìn trang), rồi CHỈ dựng
    widget + raster cho một cửa sổ nhỏ các trang quanh vùng đang xem
    (`_BUFFER_PAGES` trang đệm mỗi phía). Cuộn tới đâu, cửa sổ dịch tới đó
    (debounce theo `_scroll_timer`) — trang cũ ngoài cửa sổ được giải
    phóng khỏi layout, phần khoảng trống nó để lại được bù bằng 2 spacer
    trên/dưới nên scrollbar vẫn phản ánh đúng tổng chiều dài tài liệu.
    """

    _PAGE_W = 480
    _BUFFER_PAGES = 4        # số trang đệm thêm ngoài viewport mỗi phía
    _CACHE_MAX = 60          # số pixmap đã decode giữ lại tối đa (LRU)
    _SCROLL_DEBOUNCE_MS = 80
    _page_ready = Signal(int, int, object, int)  # (thế hệ, idx, QImage, chiều cao thật)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWidgetResizable(True)
        self.setFrameShape(QFrame.Shape.NoFrame)
        self.setStyleSheet("QScrollArea { background: #ECEBE4; border: none; }")

        self._container = QWidget()
        self._container.setStyleSheet("background: #ECEBE4;")
        self._vbox = QVBoxLayout(self._container)
        self._vbox.setContentsMargins(18, 18, 18, 18)
        self._vbox.setSpacing(14)
        self._top_spacer = QWidget()
        self._top_spacer.setStyleSheet("background: transparent;")
        self._top_spacer.setFixedHeight(0)
        self._bottom_spacer = QWidget()
        self._bottom_spacer.setStyleSheet("background: transparent;")
        self._bottom_spacer.setFixedHeight(0)
        self._vbox.addWidget(self._top_spacer)
        self._vbox.addWidget(self._bottom_spacer)
        self.setWidget(self._container)

        self._source = None
        self._page_heights: list[int] = []
        self._cum_offsets: list[int] = []
        self._active_labels: dict[int, QLabel] = {}
        self._pixmap_cache: OrderedDict[int, QPixmap] = OrderedDict()
        self._window = (0, 0)
        self._load_gen = 0

        self._page_ready.connect(self._on_page_ready)
        self._scroll_timer = QTimer(self)
        self._scroll_timer.setSingleShot(True)
        self._scroll_timer.setInterval(self._SCROLL_DEBOUNCE_MS)
        self._scroll_timer.timeout.connect(self._apply_window)
        self.verticalScrollBar().valueChanged.connect(lambda _v: self._scroll_timer.start())

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._scroll_timer.start()

    # ── API công khai ────────────────────────────────────────────────────────

    def show_message(self, msg: str):
        self._reset(close_source=True)
        lbl = QLabel(msg)
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl.setWordWrap(True)
        lbl.setStyleSheet("color: #8A8A85; font-size: 13px; background: transparent;")
        pos = self._vbox.indexOf(self._top_spacer) + 1
        self._vbox.insertWidget(pos, lbl)

    def load_source(self, source):
        """`source`: có .page_count, .page_size(i), .png_path(i), .close()."""
        self._load_gen += 1
        gen = self._load_gen
        self._reset(close_source=True)
        self._source = source

        page_count = source.page_count
        if page_count == 0:
            self.show_message("Không có trang nào.")
            return

        w0, h0 = source.page_size(0)
        est_h = max(1, int(h0 * (self._PAGE_W / w0)))
        self._page_heights = [est_h] * page_count
        self._recompute_offsets()
        self._bottom_spacer.setFixedHeight(self._span_height(0, page_count))
        self.verticalScrollBar().setValue(0)
        self._window = (0, 0)
        self._apply_window()

    def cleanup(self):
        self._reset(close_source=True)

    def goto_page(self, idx: int):
        """Cuộn tới trang `idx` (0-based) — dùng cho kết quả tìm kiếm từ khóa."""
        if not self._cum_offsets or not 0 <= idx < len(self._cum_offsets):
            return
        self.verticalScrollBar().setValue(self._cum_offsets[idx])

    # ── Bố cục ảo hoá ────────────────────────────────────────────────────────

    def _reset(self, close_source: bool):
        pos = self._vbox.indexOf(self._top_spacer) + 1
        end = self._vbox.indexOf(self._bottom_spacer)
        while end > pos:
            item = self._vbox.takeAt(pos)
            if item.widget():
                item.widget().deleteLater()
            end -= 1
        self._active_labels = {}
        self._pixmap_cache = OrderedDict()
        self._page_heights = []
        self._cum_offsets = []
        self._window = (0, 0)
        self._top_spacer.setFixedHeight(0)
        self._bottom_spacer.setFixedHeight(0)
        if close_source and self._source is not None:
            self._source.close()
        self._source = None

    def _recompute_offsets(self):
        offsets = []
        offset_y = 0
        gap = self._vbox.spacing()
        for height in self._page_heights:
            offsets.append(offset_y)
            offset_y += height + gap
        self._cum_offsets = offsets

    def _span_height(self, start: int, end: int) -> int:
        """Chiều cao (kèm khoảng cách nội bộ) của các trang [start, end)."""
        if end <= start:
            return 0
        gap = self._vbox.spacing()
        return sum(self._page_heights[start:end]) + gap * (end - start - 1)

    def _index_at_offset(self, offset_y: int) -> int:
        if not self._cum_offsets:
            return 0
        idx = bisect.bisect_right(self._cum_offsets, offset_y) - 1
        return max(0, min(idx, len(self._cum_offsets) - 1))

    def _desired_window(self) -> tuple[int, int]:
        page_count = len(self._page_heights)
        if page_count == 0:
            return (0, 0)
        top_y = self.verticalScrollBar().value()
        bottom_y = top_y + max(self.viewport().height(), 1)
        start = self._index_at_offset(top_y)
        end = self._index_at_offset(bottom_y) + 1
        start = max(0, start - self._BUFFER_PAGES)
        end = min(page_count, end + self._BUFFER_PAGES)
        return start, end

    def _apply_window(self):
        if self._source is None or not self._page_heights:
            return
        gen = self._load_gen
        start, end = self._desired_window()
        if (start, end) == self._window:
            return
        self._window = (start, end)

        for lbl in self._active_labels.values():
            self._vbox.removeWidget(lbl)
            lbl.deleteLater()
        self._active_labels = {}

        self._top_spacer.setFixedHeight(self._span_height(0, start))
        self._bottom_spacer.setFixedHeight(self._span_height(end, len(self._page_heights)))

        insert_pos = self._vbox.indexOf(self._top_spacer) + 1
        to_decode: list[int] = []
        for offset, idx in enumerate(range(start, end)):
            lbl = QLabel()
            lbl.setFixedSize(QSize(self._PAGE_W, self._page_heights[idx]))
            lbl.setStyleSheet(
                "background: white; border: 1px solid #C9C8C0; border-radius: 3px;")
            cached = self._pixmap_cache.get(idx)
            if cached is not None:
                self._pixmap_cache.move_to_end(idx)
                lbl.setPixmap(cached)
            else:
                to_decode.append(idx)
            self._vbox.insertWidget(insert_pos + offset, lbl,
                                    alignment=Qt.AlignmentFlag.AlignHCenter)
            self._active_labels[idx] = lbl

        if to_decode:
            source = self._source
            threading.Thread(
                target=self._decode_pages, args=(gen, source, to_decode), daemon=True,
            ).start()

    # ── Raster + decode nền ─────────────────────────────────────────────────

    def _decode_pages(self, gen: int, source, indices: list[int]):
        from PIL import Image
        for idx in indices:
            if gen != self._load_gen:
                return   # đã load file khác hoặc cửa sổ đã dịch — dừng, khỏi phí công
            try:
                png_path = source.png_path(idx)
                with Image.open(png_path) as img:
                    new_h = int(img.height * (self._PAGE_W / img.width))
                    resized = img.resize((self._PAGE_W, new_h), Image.LANCZOS)
                    qimg = pil_to_qimage(resized)
            except Exception:
                continue
            try:
                self._page_ready.emit(gen, idx, qimg, new_h)
            except RuntimeError:
                return   # widget đã bị hủy (đóng app giữa chừng lúc còn decode)

    def _on_page_ready(self, gen: int, idx: int, qimage, real_h: int):
        if gen != self._load_gen:
            return
        pixmap = QPixmap.fromImage(qimage)
        self._pixmap_cache[idx] = pixmap
        self._pixmap_cache.move_to_end(idx)
        while len(self._pixmap_cache) > self._CACHE_MAX:
            self._pixmap_cache.popitem(last=False)

        if idx < len(self._page_heights) and real_h != self._page_heights[idx]:
            self._page_heights[idx] = real_h
            self._recompute_offsets()
            start, end = self._window
            self._top_spacer.setFixedHeight(self._span_height(0, start))
            self._bottom_spacer.setFixedHeight(self._span_height(end, len(self._page_heights)))

        lbl = self._active_labels.get(idx)
        if lbl is not None:
            lbl.setFixedSize(QSize(self._PAGE_W, real_h))
            lbl.setPixmap(pixmap)


class PreviewPanel(QFrame):
    convert_requested = Signal()
    extract_text_requested = Signal()
    _preview_ready = Signal(int, object)   # (thế hệ, source)
    _preview_err = Signal(int, str)
    _pdf_text_ready = Signal(int, list)    # (thế hệ, văn bản từng trang)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("previewPanel")
        self.setMinimumWidth(380)
        self._gen = 0   # token thế hệ — loại kết quả render cũ về muộn

        lay = QVBoxLayout(self)
        lay.setContentsMargins(1, 1, 1, 1)
        lay.setSpacing(0)

        # ── Header ────────────────────────────────────────────────────────────
        hdr = QFrame()
        hdr.setFixedHeight(46)
        hdr_lay = QHBoxLayout(hdr)
        hdr_lay.setContentsMargins(14, 0, 10, 0)
        hdr_lay.setSpacing(8)

        self.badge = QLabel("PDF")
        self.badge.setObjectName("previewBadge")
        hdr_lay.addWidget(self.badge)

        self.title = QLabel("")
        self.title.setObjectName("previewTitle")
        hdr_lay.addWidget(self.title, stretch=1)

        self._current_path: str | None = None

        self.open_folder_btn = QPushButton("Mở thư mục")
        self.open_folder_btn.setObjectName("previewOpenBtn")
        self.open_folder_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.open_folder_btn.setVisible(False)
        self.open_folder_btn.clicked.connect(self._open_folder)
        hdr_lay.addWidget(self.open_folder_btn)

        self.open_with_btn = QPushButton("Mở bằng…")
        self.open_with_btn.setObjectName("previewOpenWithBtn")
        self.open_with_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.open_with_btn.setVisible(False)
        self.open_with_btn.clicked.connect(self._open_with)
        hdr_lay.addWidget(self.open_with_btn)

        self.extract_text_btn = QPushButton("Trích xuất văn bản")
        self.extract_text_btn.setObjectName("previewOpenBtn")
        self.extract_text_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.extract_text_btn.setVisible(False)
        self.extract_text_btn.clicked.connect(self.extract_text_requested.emit)
        hdr_lay.addWidget(self.extract_text_btn)

        convert_btn = QPushButton("Chuyển đổi định dạng")
        convert_btn.setObjectName("convertBtn")
        convert_btn.clicked.connect(self.convert_requested.emit)
        hdr_lay.addWidget(convert_btn)
        lay.addWidget(hdr)

        # ── Tìm kiếm từ khóa (chỉ hiện với PDF) ──────────────────────────────
        self._search_bar = QFrame()
        self._search_bar.setVisible(False)
        search_lay = QHBoxLayout(self._search_bar)
        search_lay.setContentsMargins(14, 6, 10, 6)
        search_lay.setSpacing(6)
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Tìm từ khóa trong PDF…")
        self.search_input.returnPressed.connect(self._search_next)
        self.search_input.textChanged.connect(self._run_search)
        search_lay.addWidget(self.search_input, stretch=1)
        self.search_count = QLabel("")
        self.search_count.setStyleSheet("color: #8A8A85; font-size: 12px;")
        search_lay.addWidget(self.search_count)
        prev_btn = QPushButton("‹")
        prev_btn.setFixedWidth(28)
        prev_btn.clicked.connect(self._search_prev)
        search_lay.addWidget(prev_btn)
        next_btn = QPushButton("›")
        next_btn.setFixedWidth(28)
        next_btn.clicked.connect(self._search_next)
        search_lay.addWidget(next_btn)
        lay.addWidget(self._search_bar)

        # ── Trang xem trước ───────────────────────────────────────────────────
        self.pages = PagePreviewWidget()
        lay.addWidget(self.pages, stretch=1)

        footer = QLabel("Xem trước tài liệu — giữ nguyên format")
        footer.setObjectName("previewFooter")
        footer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(footer)

        self._pdf_pages_text: list[str] = []
        self._search_matches: list[int] = []
        self._search_pos = -1

        self._preview_ready.connect(self._on_ready)
        self._preview_err.connect(self._on_err)
        self._pdf_text_ready.connect(self._on_pdf_text_ready)

    # ── Nạp file (mở nguồn trang ở luồng nền) ───────────────────────────────

    def load_file(self, path: str, file_type: str, name: str):
        self._gen += 1
        gen = self._gen
        self.badge.setText(FILE_BADGE.get(file_type, "FILE"))
        self.title.setText(name)

        self._current_path = path
        self.open_folder_btn.setVisible(True)
        self.open_with_btn.setVisible(True)
        self.open_with_btn.setText(f"Mở bằng {OPEN_WITH_LABEL.get(file_type, 'ứng dụng')}")

        is_pdf = file_type == "pdf"
        self.extract_text_btn.setVisible(is_pdf)
        self._search_bar.setVisible(is_pdf)
        self.search_input.clear()
        self._pdf_pages_text = []
        self._search_matches = []
        self._search_pos = -1
        self.search_count.setText("")
        if is_pdf and HAS_PDF_TEXT:
            def _load_text():
                try:
                    pages = pdf_page_texts(path)
                    self._pdf_text_ready.emit(gen, pages)
                except Exception:
                    pass
            threading.Thread(target=_load_text, daemon=True).start()

        factory = _SOURCE_FACTORIES.get(file_type)
        if factory is None:
            self.pages.show_message("Định dạng chưa hỗ trợ xem trước.")
            return
        if file_type not in ("image",) and not HAS_PAGEPREVIEW:
            self.pages.show_message(
                "Cần cài PyMuPDF + Pillow để xem trước tài liệu.")
            return

        self.pages.show_message("Đang mở tài liệu…")

        def _work():
            try:
                source = factory(path)
                self._preview_ready.emit(gen, source)
            except Exception as exc:
                self._preview_err.emit(gen, str(exc))

        threading.Thread(target=_work, daemon=True).start()

    def _on_ready(self, gen: int, source):
        if gen != self._gen:
            source.close()   # kết quả của file đã bị thay — bỏ, đóng nguồn
            return
        self.pages.load_source(source)

    def _on_err(self, gen: int, msg: str):
        if gen != self._gen:
            return
        self.pages.show_message(f"Không xem trước được:\n{msg}")

    def cleanup(self):
        self.pages.cleanup()

    # ── Tìm kiếm từ khóa trong PDF ───────────────────────────────────────────

    def _on_pdf_text_ready(self, gen: int, pages: list):
        if gen != self._gen:
            return
        self._pdf_pages_text = pages
        if self.search_input.text().strip():
            self._run_search()

    def _run_search(self):
        query = self.search_input.text().strip().lower()
        if not query or not self._pdf_pages_text:
            self._search_matches = []
            self._search_pos = -1
            self.search_count.setText("")
            return
        self._search_matches = [
            page_index for page_index, page_text in enumerate(self._pdf_pages_text) if query in page_text.lower()]
        self._search_pos = 0 if self._search_matches else -1
        self._update_search_count()
        if self._search_matches:
            self.pages.goto_page(self._search_matches[0])

    def _update_search_count(self):
        if not self._search_matches:
            self.search_count.setText(
                "Không tìm thấy" if self.search_input.text().strip() else "")
            return
        page = self._search_matches[self._search_pos] + 1
        self.search_count.setText(
            f"Trang {page} ({self._search_pos + 1}/{len(self._search_matches)})")

    def _search_next(self):
        if not self._search_matches:
            self._run_search()
            return
        self._search_pos = (self._search_pos + 1) % len(self._search_matches)
        self._update_search_count()
        self.pages.goto_page(self._search_matches[self._search_pos])

    def _search_prev(self):
        if not self._search_matches:
            return
        self._search_pos = (self._search_pos - 1) % len(self._search_matches)
        self._update_search_count()
        self.pages.goto_page(self._search_matches[self._search_pos])

    # ── 7.6: mở thư mục chứa file / mở bằng ứng dụng mặc định ────────────────

    def _open_folder(self):
        if self._current_path:
            _open_path(str(Path(self._current_path).parent))

    def _open_with(self):
        if self._current_path:
            _open_path(self._current_path)
