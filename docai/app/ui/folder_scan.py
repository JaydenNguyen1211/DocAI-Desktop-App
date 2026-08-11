"""Scan the working folder (Folder mode) — list supported files by type,
grouped by top-level subfolder, skipping unsupported formats."""
import os
from pathlib import Path

from PySide6.QtCore import QThread, Signal

from ..constants import FOLDER_EXT_MAP

from ...logging_config import get_logger, log_call

logger = get_logger(__name__)


class ScannedFile:
    """A supported file found in the working folder."""

    __slots__ = ("path", "name", "group", "ext_type")

    @log_call
    def __init__(self, path: str, name: str, group: str, ext_type: str):
        self.path = path        # absolute path
        self.name = name        # file name
        self.group = group      # first-level subfolder relative to the root ("" = at the root)
        self.ext_type = ext_type


class FolderScanWorker(QThread):
    """Recursively scans `root` on a background thread, emitting progress
    for each file examined."""

    # (files examined so far, total files in the tree, counts by supported type)
    progress = Signal(int, int, dict)
    # (list of supported ScannedFile, counts by type)
    done = Signal(list, dict)

    @log_call
    def __init__(self, root: str, parent=None):
        super().__init__(parent)
        self._root = root

    @log_call
    def run(self):
        root = Path(self._root)
        try:
            all_paths = []
            for dirpath, dirnames, filenames in os.walk(root):
                # Skip DocAI's internal trash folder (see
                # `modules.common.folder_ops.delete_file_with_undo`) — a file
                # that was just "deleted" (undoable) shouldn't show up again
                # as if it were a real file.
                dirnames[:] = [d for d in dirnames if d != ".docai_trash"]
                all_paths.extend(Path(dirpath) / fn for fn in filenames)
        except OSError:
            logger.exception("Failed to walk folder tree at %s — reporting 0 files.", root)
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
