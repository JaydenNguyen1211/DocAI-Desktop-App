"""Convert files to PDF via the LibreOffice CLI — used on macOS/Linux instead of COM."""
import os
import shutil
import subprocess
import tempfile

from ...logging_config import get_logger, log_call
from ...strings import Office as S

logger = get_logger(__name__)


class SofficeError(Exception):
    pass


@log_call
def _find_soffice() -> str:
    candidates = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        shutil.which("libreoffice"),
        shutil.which("soffice"),
    ]
    for c in candidates:
        if c and os.path.isfile(c):
            return c
    raise SofficeError(S.LIBREOFFICE_NOT_FOUND)


@log_call
def convert_to_pdf(input_path: str, out_path: str) -> str:
    """Convert the file to PDF then move the result to `out_path`."""
    soffice = _find_soffice()
    abs_input = os.path.abspath(input_path)
    abs_out = os.path.abspath(out_path)

    with tempfile.TemporaryDirectory(prefix="docai_lo_") as tmp_dir:
        proc = subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp_dir, abs_input],
            capture_output=True, text=True, timeout=60,
        )
        if proc.returncode != 0:
            detail = (proc.stderr or proc.stdout).strip()
            raise SofficeError(S.LIBREOFFICE_ERROR.format(detail=detail))

        base = os.path.splitext(os.path.basename(abs_input))[0]
        tmp_pdf = os.path.join(tmp_dir, base + ".pdf")
        if not os.path.exists(tmp_pdf):
            raise SofficeError(S.LIBREOFFICE_NO_PDF)

        out_dir = os.path.dirname(abs_out)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        shutil.move(tmp_pdf, abs_out)

    return abs_out
