""""Save file" dialog for output files still in staging (temporary) form —
used after `modules.common.doc_set.process_document_set()` returns
`output_files`. There's no dedicated design for this step yet, so it uses
the OS's standard save-file dialog, the same way
`main_window._on_extract_text()` already does for the PDF text-extraction
feature.

The file is only actually created at the location the user picks AFTER
confirming here — the temporary staging file doesn't disappear on its own,
it's just copied to the save location.
"""
import shutil
from pathlib import Path

from PySide6.QtWidgets import QFileDialog, QWidget

from ...logging_config import get_logger, log_call
from ...strings import SaveOutput as S

logger = get_logger(__name__)

_FILTER_BY_EXT = {
    ".docx": S.FILTER_DOCX,
    ".xlsx": S.FILTER_XLSX,
    ".pdf": S.FILTER_PDF,
    ".pptx": S.FILTER_PPTX,
}


@log_call
def save_staged_file(parent: QWidget, staged_path: str, suggested_name: str = "") -> str | None:
    """Open a save-location dialog and copy the staged file there. Returns the
    saved path, or None if the user canceled / a write error occurred."""
    staged = Path(staged_path)
    default_name = suggested_name or staged.name
    file_filter = _FILTER_BY_EXT.get(staged.suffix.lower(), S.FILTER_ALL)

    out_path, _ = QFileDialog.getSaveFileName(
        parent, S.DIALOG_TITLE, default_name, file_filter)
    if not out_path:
        return None
    try:
        shutil.copyfile(staged_path, out_path)
    except OSError:
        logger.exception("Failed to copy staged file %s -> %s", staged_path, out_path)
        return None
    return out_path
