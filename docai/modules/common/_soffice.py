"""Chuyển đổi file sang PDF qua LibreOffice CLI — dùng trên macOS/Linux thay COM."""
import os
import shutil
import subprocess
import tempfile


class SofficeError(Exception):
    pass


def _find_soffice() -> str:
    candidates = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        shutil.which("libreoffice"),
        shutil.which("soffice"),
    ]
    for c in candidates:
        if c and os.path.isfile(c):
            return c
    raise SofficeError(
        "Không tìm thấy LibreOffice. Hãy cài LibreOffice (https://libreoffice.org) "
        "hoặc Microsoft Office để chuyển đổi / xem trước file Word, PowerPoint, Excel."
    )


def convert_to_pdf(input_path: str, out_path: str) -> str:
    """Chuyển đổi file sang PDF rồi di chuyển kết quả tới `out_path`."""
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
            raise SofficeError(f"LibreOffice báo lỗi: {detail}")

        base = os.path.splitext(os.path.basename(abs_input))[0]
        tmp_pdf = os.path.join(tmp_dir, base + ".pdf")
        if not os.path.exists(tmp_pdf):
            raise SofficeError("LibreOffice không tạo được file PDF.")

        out_dir = os.path.dirname(abs_out)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        shutil.move(tmp_pdf, abs_out)

    return abs_out
