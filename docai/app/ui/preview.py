"""Document preview panel — renders pages via Word COM / LibreOffice / fitz / PIL."""
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

from ...logging_config import get_logger, log_call
from ...strings import Preview as S

logger = get_logger(__name__)

_SOURCE_FACTORIES = {
    "word": word_page_source,
    "excel": excel_page_source,
    "pdf": pdf_page_source,
    "image": image_page_source,
    "ppt": ppt_page_source,
}


@log_call
def _open_path(path: str) -> None:
    if sys.platform == "win32":
        os.startfile(path)  # noqa: S606
    elif sys.platform == "darwin":
        subprocess.run(["open", path], check=False)
    else:
        subprocess.run(["xdg-open", path], check=False)


class PagePreviewWidget(QScrollArea):
    """List of preview pages — virtualized against the viewport.

    For a document with hundreds to thousands of pages, RASTERIZING each
    page and CREATING A WIDGET for each page both take time proportional to
    the page count — doing it all right when the file opens (whether on a
    background thread or the main thread) makes the user wait a long time
    before seeing anything, or freezes the UI if run synchronously on the
    main thread.

    This approach instead: opens the document to get the page count + first
    page size (a metadata read, no rasterizing — cheap even for a document
    with thousands of pages), then ONLY builds widgets + rasterizes for a
    small window of pages around the currently visible area
    (`_BUFFER_PAGES` buffer pages on each side). As the user scrolls, the
    window shifts (debounced via `_scroll_timer`) — pages that fall outside
    the window are released from the layout, and the gap they leave behind
    is compensated by 2 spacers (top/bottom) so the scrollbar still
    reflects the document's true total length.
    """

    _PAGE_W = 480
    _BUFFER_PAGES = 4        # extra buffer pages outside the viewport on each side
    _CACHE_MAX = 60          # max decoded pixmaps kept around (LRU)
    _SCROLL_DEBOUNCE_MS = 80
    _page_ready = Signal(int, int, object, int)  # (generation, idx, QImage, real height)

    @log_call
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

    # ── Public API ───────────────────────────────────────────────────────────

    @log_call
    def show_message(self, msg: str):
        self._reset(close_source=True)
        lbl = QLabel(msg)
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl.setWordWrap(True)
        lbl.setStyleSheet("color: #8A8A85; font-size: 13px; background: transparent;")
        pos = self._vbox.indexOf(self._top_spacer) + 1
        self._vbox.insertWidget(pos, lbl)

    @log_call
    def load_source(self, source):
        """`source`: has .page_count, .page_size(i), .png_path(i), .close()."""
        self._load_gen += 1
        gen = self._load_gen
        self._reset(close_source=True)
        self._source = source

        page_count = source.page_count
        if page_count == 0:
            self.show_message(S.NO_PAGES)
            return

        w0, h0 = source.page_size(0)
        est_h = max(1, int(h0 * (self._PAGE_W / w0)))
        self._page_heights = [est_h] * page_count
        self._recompute_offsets()
        self._bottom_spacer.setFixedHeight(self._span_height(0, page_count))
        self.verticalScrollBar().setValue(0)
        self._window = (0, 0)
        self._apply_window()

    @log_call
    def cleanup(self):
        self._reset(close_source=True)

    @log_call
    def goto_page(self, idx: int):
        """Scroll to page `idx` (0-based) — used for keyword search results."""
        if not self._cum_offsets or not 0 <= idx < len(self._cum_offsets):
            return
        self.verticalScrollBar().setValue(self._cum_offsets[idx])

    # ── Virtualized layout ───────────────────────────────────────────────────

    @log_call
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

    @log_call
    def _recompute_offsets(self):
        offsets = []
        offset_y = 0
        gap = self._vbox.spacing()
        for height in self._page_heights:
            offsets.append(offset_y)
            offset_y += height + gap
        self._cum_offsets = offsets

    @log_call
    def _span_height(self, start: int, end: int) -> int:
        """Height (including inter-page spacing) of pages [start, end)."""
        if end <= start:
            return 0
        gap = self._vbox.spacing()
        return sum(self._page_heights[start:end]) + gap * (end - start - 1)

    @log_call
    def _index_at_offset(self, offset_y: int) -> int:
        if not self._cum_offsets:
            return 0
        idx = bisect.bisect_right(self._cum_offsets, offset_y) - 1
        return max(0, min(idx, len(self._cum_offsets) - 1))

    @log_call
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

    @log_call
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

    # ── Background rasterizing + decoding ───────────────────────────────────

    @log_call
    def _decode_pages(self, gen: int, source, indices: list[int]):
        from PIL import Image
        for idx in indices:
            if gen != self._load_gen:
                return   # a different file was loaded or the window moved — stop, no point continuing
            try:
                png_path = source.png_path(idx)
                with Image.open(png_path) as img:
                    new_h = int(img.height * (self._PAGE_W / img.width))
                    resized = img.resize((self._PAGE_W, new_h), Image.LANCZOS)
                    qimg = pil_to_qimage(resized)
            except Exception:
                logger.warning("Failed to decode preview page %d — skipping.", idx,
                               exc_info=True)
                continue
            try:
                self._page_ready.emit(gen, idx, qimg, new_h)
            except RuntimeError:
                return   # widget was already destroyed (app closed mid-decode)

    @log_call
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
    _preview_ready = Signal(int, object)   # (generation, source)
    _preview_err = Signal(int, str)
    _pdf_text_ready = Signal(int, list)    # (generation, per-page text)

    @log_call
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("previewPanel")
        self.setMinimumWidth(380)
        self._gen = 0   # generation token — discards stale render results that arrive late

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

        self.open_folder_btn = QPushButton(S.OPEN_FOLDER)
        self.open_folder_btn.setObjectName("previewOpenBtn")
        self.open_folder_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.open_folder_btn.setVisible(False)
        self.open_folder_btn.clicked.connect(self._open_folder)
        hdr_lay.addWidget(self.open_folder_btn)

        self.open_with_btn = QPushButton(S.OPEN_WITH)
        self.open_with_btn.setObjectName("previewOpenWithBtn")
        self.open_with_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.open_with_btn.setVisible(False)
        self.open_with_btn.clicked.connect(self._open_with)
        hdr_lay.addWidget(self.open_with_btn)

        self.extract_text_btn = QPushButton(S.EXTRACT_TEXT)
        self.extract_text_btn.setObjectName("previewOpenBtn")
        self.extract_text_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.extract_text_btn.setVisible(False)
        self.extract_text_btn.clicked.connect(self.extract_text_requested.emit)
        hdr_lay.addWidget(self.extract_text_btn)

        convert_btn = QPushButton(S.CONVERT)
        convert_btn.setObjectName("convertBtn")
        convert_btn.clicked.connect(self.convert_requested.emit)
        hdr_lay.addWidget(convert_btn)
        lay.addWidget(hdr)

        # ── Keyword search (only shown for PDF) ─────────────────────────────
        self._search_bar = QFrame()
        self._search_bar.setVisible(False)
        search_lay = QHBoxLayout(self._search_bar)
        search_lay.setContentsMargins(14, 6, 10, 6)
        search_lay.setSpacing(6)
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText(S.SEARCH_PLACEHOLDER)
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

        # ── Preview pages ────────────────────────────────────────────────────
        self.pages = PagePreviewWidget()
        lay.addWidget(self.pages, stretch=1)

        footer = QLabel(S.FOOTER)
        footer.setObjectName("previewFooter")
        footer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(footer)

        self._pdf_pages_text: list[str] = []
        self._search_matches: list[int] = []
        self._search_pos = -1

        self._preview_ready.connect(self._on_ready)
        self._preview_err.connect(self._on_err)
        self._pdf_text_ready.connect(self._on_pdf_text_ready)

    # ── Load a file (open the page source on a background thread) ──────────

    @log_call
    def load_file(self, path: str, file_type: str, name: str):
        self._gen += 1
        gen = self._gen
        self.badge.setText(FILE_BADGE.get(file_type, "FILE"))
        self.title.setText(name)

        self._current_path = path
        self.open_folder_btn.setVisible(True)
        self.open_with_btn.setVisible(True)
        self.open_with_btn.setText(
            S.OPEN_WITH_TYPE.format(label=OPEN_WITH_LABEL.get(file_type, S.OPEN_WITH_DEFAULT_LABEL)))

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
                    logger.warning("Failed to load PDF text for search: %s", path,
                                   exc_info=True)
            threading.Thread(target=_load_text, daemon=True).start()

        factory = _SOURCE_FACTORIES.get(file_type)
        if factory is None:
            self.pages.show_message(S.UNSUPPORTED_FORMAT)
            return
        if file_type not in ("image",) and not HAS_PAGEPREVIEW:
            self.pages.show_message(S.NEEDS_DEPENDENCIES)
            return

        self.pages.show_message(S.OPENING)

        def _work():
            try:
                source = factory(path)
                self._preview_ready.emit(gen, source)
            except Exception as exc:
                logger.exception("Failed to build preview page source for %s", path)
                self._preview_err.emit(gen, str(exc))

        threading.Thread(target=_work, daemon=True).start()

    @log_call
    def _on_ready(self, gen: int, source):
        if gen != self._gen:
            source.close()   # result for a file that's since been replaced — discard, close the source
            return
        self.pages.load_source(source)

    @log_call
    def _on_err(self, gen: int, msg: str):
        if gen != self._gen:
            return
        self.pages.show_message(S.PREVIEW_FAILED.format(error=msg))

    @log_call
    def cleanup(self):
        self.pages.cleanup()

    # ── Keyword search inside a PDF ──────────────────────────────────────────

    @log_call
    def _on_pdf_text_ready(self, gen: int, pages: list):
        if gen != self._gen:
            return
        self._pdf_pages_text = pages
        if self.search_input.text().strip():
            self._run_search()

    @log_call
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

    @log_call
    def _update_search_count(self):
        if not self._search_matches:
            self.search_count.setText(
                S.NOT_FOUND if self.search_input.text().strip() else "")
            return
        page = self._search_matches[self._search_pos] + 1
        self.search_count.setText(
            S.SEARCH_RESULT.format(page=page, pos=self._search_pos + 1, total=len(self._search_matches)))

    @log_call
    def _search_next(self):
        if not self._search_matches:
            self._run_search()
            return
        self._search_pos = (self._search_pos + 1) % len(self._search_matches)
        self._update_search_count()
        self.pages.goto_page(self._search_matches[self._search_pos])

    @log_call
    def _search_prev(self):
        if not self._search_matches:
            return
        self._search_pos = (self._search_pos - 1) % len(self._search_matches)
        self._update_search_count()
        self.pages.goto_page(self._search_matches[self._search_pos])

    # ── 7.6: open the containing folder / open with the default app ────────

    @log_call
    def _open_folder(self):
        if self._current_path:
            _open_path(str(Path(self._current_path).parent))

    @log_call
    def _open_with(self):
        if self._current_path:
            _open_path(self._current_path)
