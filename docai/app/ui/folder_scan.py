"""Quét thư mục làm việc (chế độ Thư mục) — liệt kê file hỗ trợ theo loại,
nhóm theo thư mục con cấp 1, bỏ qua định dạng chưa hỗ trợ."""
import os
from pathlib import Path

from PySide6.QtCore import QThread, Signal

from ..constants import FOLDER_EXT_MAP


class ScannedFile:
    """Một file hỗ trợ tìm thấy trong thư mục làm việc."""

    __slots__ = ("path", "name", "group", "ext_type")

    def __init__(self, path: str, name: str, group: str, ext_type: str):
        self.path = path        # đường dẫn tuyệt đối
        self.name = name        # tên file
        self.group = group      # thư mục con cấp 1 so với gốc ("" = ngay tại gốc)
        self.ext_type = ext_type


class FolderScanWorker(QThread):
    """Quét đệ quy `root` ở luồng nền, phát tiến trình theo từng file đã xét."""

    # (số file đã xét, tổng số file trong cây, đếm file hỗ trợ theo loại)
    progress = Signal(int, int, dict)
    # (danh sách ScannedFile hỗ trợ, đếm theo loại)
    done = Signal(list, dict)

    def __init__(self, root: str, parent=None):
        super().__init__(parent)
        self._root = root

    def run(self):
        root = Path(self._root)
        try:
            all_paths = [
                Path(dirpath) / fn
                for dirpath, _dirnames, filenames in os.walk(root)
                for fn in filenames
            ]
        except OSError:
            all_paths = []

        total = len(all_paths)
        results: list[ScannedFile] = []
        counts: dict[str, int] = {}

        for scanned_index, path in enumerate(all_paths, start=1):
            ext_type = FOLDER_EXT_MAP.get(path.suffix.lower())
            if ext_type is not None:
                try:
                    rel_parts = path.relative_to(root).parts
                except ValueError:
                    rel_parts = (path.name,)
                group = rel_parts[0] if len(rel_parts) > 1 else ""
                results.append(ScannedFile(str(path), path.name, group, ext_type))
                counts[ext_type] = counts.get(ext_type, 0) + 1
            self.progress.emit(scanned_index, total, dict(counts))

        self.done.emit(results, counts)
