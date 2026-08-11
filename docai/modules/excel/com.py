"""Export PNG images of Excel charts via COM (Excel.Application) to compose
into the self-drawn preview — only works on Windows. On macOS/Linux,
`export_all_charts` returns {} and charts show as a placeholder;
`export_workbook_pdf` uses LibreOffice instead.

Uses exactly 1 narrow COM API — `Chart.Export()` — does NOT touch the
Application window's position/size. Opens the file READ-ONLY and always
`Close(SaveChanges=False)` — so Excel never recalculates/"fixes" and
overwrites XML in ways the openpyxl code doesn't anticipate.
"""
import os
import sys

from ...logging_config import get_logger, log_call
from ...strings import ExcelCom as S

logger = get_logger(__name__)


class ExcelComError(Exception):
    """Couldn't export the PDF — message shown to the user in Vietnamese."""


@log_call
def export_workbook_pdf(path: str, out_path: str) -> str:
    """Export the entire workbook to PDF.

    Windows: uses Excel COM (`ExportAsFixedFormat`).
    macOS/Linux: uses the LibreOffice CLI.
    """
    if sys.platform == "darwin":
        from ..common._msoffice_mac import is_available, excel_to_pdf as ms_excel_to_pdf, MsOfficeMacError
        from ..common._soffice import convert_to_pdf, SofficeError
        if is_available("excel"):
            try:
                ms_excel_to_pdf(os.path.abspath(path), os.path.abspath(out_path))
                return os.path.abspath(out_path)
            except MsOfficeMacError as exc:
                raise ExcelComError(str(exc))
        try:
            return convert_to_pdf(os.path.abspath(path), os.path.abspath(out_path))
        except SofficeError as exc:
            raise ExcelComError(str(exc))

    if sys.platform != "win32":
        from ..common._soffice import convert_to_pdf, SofficeError
        try:
            return convert_to_pdf(os.path.abspath(path), os.path.abspath(out_path))
        except SofficeError as exc:
            raise ExcelComError(str(exc))

    try:
        import pythoncom
        import win32com.client
    except ImportError:
        raise ExcelComError(S.MISSING_EXCEL)

    abs_path = os.path.abspath(path)
    abs_out = os.path.abspath(out_path)
    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(abs_path, ReadOnly=True, UpdateLinks=0)
        wb.ExportAsFixedFormat(0, abs_out)  # 0 = xlTypePDF
    except Exception as exc:
        raise ExcelComError(S.EXPORT_FAILED.format(error=exc))
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            logger.debug("Failed to close Excel workbook cleanly after PDF export.",
                         exc_info=True)
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            logger.debug("Failed to quit Excel.Application cleanly after PDF export.",
                         exc_info=True)
        pythoncom.CoUninitialize()
    return abs_out


@log_call
def export_all_charts(path: str, tmp_dir: str) -> dict:
    """Returns {(sheet_name, 0_based_chart_index): png_path}.

    Only works on Windows (Excel COM). On macOS/Linux returns {} so the
    caller draws its own placeholder.
    """
    if sys.platform != "win32":
        return {}

    try:
        import pythoncom
        import win32com.client
    except ImportError:
        logger.debug("pywin32 not available — export_all_charts() returns {} (placeholders used).")
        return {}

    abs_path = os.path.abspath(path)
    result: dict = {}
    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(abs_path, ReadOnly=True, UpdateLinks=0)
        for sheet in wb.Sheets:
            try:
                chart_objects = sheet.ChartObjects()
                count = chart_objects.Count
            except Exception:
                logger.warning("Could not enumerate chart objects on sheet %r — skipping.",
                               getattr(sheet, "Name", "?"), exc_info=True)
                continue
            for chart_number in range(1, count + 1):
                try:
                    chart = chart_objects.Item(chart_number).Chart
                    safe_name = "".join(char if char.isalnum() else "_" for char in str(sheet.Name))
                    png_path = os.path.join(tmp_dir, f"chart_{safe_name}_{chart_number}.png")
                    chart.Export(png_path)
                    result[(sheet.Name, chart_number - 1)] = png_path
                except Exception:
                    logger.warning("Failed to export chart #%d on sheet %r — skipping.",
                                   chart_number, getattr(sheet, "Name", "?"), exc_info=True)
                    continue
    except Exception:
        logger.exception("export_all_charts failed for %s — returning {} (placeholders used).",
                         abs_path)
        return {}
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            logger.debug("Failed to close Excel workbook cleanly after chart export.",
                         exc_info=True)
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            logger.debug("Failed to quit Excel.Application cleanly after chart export.",
                         exc_info=True)
        pythoncom.CoUninitialize()
    return result
