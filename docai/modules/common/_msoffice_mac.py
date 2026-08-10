"""Convert Office files to PDF via Microsoft Office for Mac (AppleScript/osascript).

On macOS, preferred over LibreOffice when Microsoft Office is installed —
gives higher-fidelity output for complex formatting and embedded charts.
"""
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

from ...logging_config import get_logger, log_call

logger = get_logger(__name__)


class MsOfficeMacError(Exception):
    pass


_APP_PATHS = {
    "word": "/Applications/Microsoft Word.app",
    "excel": "/Applications/Microsoft Excel.app",
    "powerpoint": "/Applications/Microsoft PowerPoint.app",
}


@log_call
def is_available(app: str) -> bool:
    """Return True if the given Office app ('word'|'excel'|'powerpoint') is installed."""
    return os.path.isdir(_APP_PATHS.get(app, ""))


@log_call
def _run_applescript(script: str, timeout: int = 90) -> None:
    proc = subprocess.run(
        ["osascript", "-e", script],
        capture_output=True, text=True, timeout=timeout,
    )
    if proc.returncode != 0:
        detail = (proc.stderr or proc.stdout).strip()
        raise MsOfficeMacError(detail or "AppleScript thất bại không rõ nguyên nhân.")


@log_call
def _safe_path(path: str) -> str:
    if '"' in path:
        raise MsOfficeMacError(
            f"Đường dẫn chứa ký tự không hợp lệ (dấu nháy kép): {path}"
        )
    return path


@log_call
def _home_tmp() -> str:
    """Return a temp dir inside the user home — accessible to Office apps (not /tmp)."""
    cache = Path.home() / "Library" / "Caches" / "docai_office_tmp"
    cache.mkdir(parents=True, exist_ok=True)
    return tempfile.mkdtemp(dir=cache)


@log_call
def _with_accessible_paths(input_path: str, out_path: str, fn) -> None:
    """Run fn(src, dst). If either path is outside the user home, stage via a
    temp dir inside ~/Library/Caches so Word/Excel/PowerPoint can access them."""
    home = str(Path.home())
    need_stage = not input_path.startswith(home) or not out_path.startswith(home)
    if not need_stage:
        fn(input_path, out_path)
        return

    tmp_dir = _home_tmp()
    try:
        src = shutil.copy2(input_path, tmp_dir)
        dst = os.path.join(tmp_dir, Path(out_path).name)
        fn(src, dst)
        out_dir = os.path.dirname(out_path)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        shutil.move(dst, out_path)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


@log_call
def word_to_pdf(input_path: str, out_path: str) -> None:
    """Convert .docx → PDF using docx2pdf (wraps Microsoft Word on macOS)."""
    try:
        from docx2pdf import convert as _d2p_convert
    except ImportError:
        raise MsOfficeMacError(
            "Cần cài docx2pdf: pip install docx2pdf"
        )
    try:
        _d2p_convert(input_path, out_path)
    except Exception as exc:
        raise MsOfficeMacError(f"docx2pdf thất bại: {exc}")


@log_call
def excel_to_pdf(input_path: str, out_path: str) -> None:
    def _convert(src: str, dst: str) -> None:
        script = (
            f'tell application "Microsoft Excel"\n'
            f'    open POSIX file "{_safe_path(src)}"\n'
            f'    delay 1\n'
            f'    set theWB to workbook 1\n'
            f'    save workbook as theWB filename "{_safe_path(dst)}" file format PDF file format\n'
            f'    close theWB saving no\n'
            f'end tell'
        )
        _run_applescript(script)
    _with_accessible_paths(input_path, out_path, _convert)


@log_call
def powerpoint_to_pdf(input_path: str, out_path: str) -> None:
    def _convert(src: str, dst: str) -> None:
        script = (
            f'tell application "Microsoft PowerPoint"\n'
            f'    open POSIX file "{_safe_path(src)}"\n'
            f'    delay 1\n'
            f'    set thePres to presentation 1\n'
            f'    save thePres in "{_safe_path(dst)}" as save as PDF\n'
            f'    close thePres saving no\n'
            f'end tell'
        )
        _run_applescript(script)
    _with_accessible_paths(input_path, out_path, _convert)
