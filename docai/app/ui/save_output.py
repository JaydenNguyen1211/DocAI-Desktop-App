"""Hộp thoại "Lưu file" cho các file output đang ở dạng staging (tạm) —
dùng sau khi `modules.common.doc_set.process_document_set()` trả về
`output_files`. Chưa có thiết kế riêng cho bước này nên dùng hộp thoại lưu
file chuẩn của hệ điều hành, theo đúng cách `main_window._on_extract_text()`
đã làm cho tính năng trích xuất văn bản PDF.

File chỉ thực sự được tạo ở vị trí người dùng chọn SAU khi xác nhận ở đây —
file staging tạm không tự động biến mất, chỉ được copy sang nơi lưu.
"""
import shutil
from pathlib import Path

from PySide6.QtWidgets import QFileDialog, QWidget

_FILTER_BY_EXT = {
    ".docx": "Word (*.docx)",
    ".xlsx": "Excel (*.xlsx)",
    ".pdf": "PDF (*.pdf)",
    ".pptx": "PowerPoint (*.pptx)",
}


def save_staged_file(parent: QWidget, staged_path: str, suggested_name: str = "") -> str | None:
    """Mở hộp thoại chọn nơi lưu, copy file staging sang đó. Trả về đường dẫn
    đã lưu, hoặc None nếu người dùng hủy / có lỗi ghi file."""
    staged = Path(staged_path)
    default_name = suggested_name or staged.name
    file_filter = _FILTER_BY_EXT.get(staged.suffix.lower(), "Tất cả file (*.*)")

    out_path, _ = QFileDialog.getSaveFileName(
        parent, "Lưu file", default_name, file_filter)
    if not out_path:
        return None
    try:
        shutil.copyfile(staged_path, out_path)
    except OSError:
        return None
    return out_path
