"""Xuất ảnh PNG của biểu đồ Excel bằng COM (Excel.Application) để ghép vào bản
xem trước tự vẽ — chỉ hoạt động trên Windows. Trên macOS/Linux, `export_all_charts`
trả về {} và biểu đồ hiện dưới dạng placeholder; `export_workbook_pdf` dùng
LibreOffice thay thế.

Chỉ dùng đúng 1 API COM hẹp — `Chart.Export()` — KHÔNG đụng đến vị trí/kích
thước cửa sổ Application. Mở file READ-ONLY và luôn `Close(SaveChanges=False)`
— không để Excel tự tính lại/"sửa" rồi ghi đè XML mà code openpyxl không
lường trước.
"""
import os
import sys


class ExcelComError(Exception):
    """Không xuất được PDF — thông báo tiếng Việt cho người dùng."""


def export_workbook_pdf(path: str, out_path: str) -> str:
    """Xuất toàn bộ workbook ra PDF.

    Windows: dùng Excel COM (`ExportAsFixedFormat`).
    macOS/Linux: dùng LibreOffice CLI.
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
        raise ExcelComError("Cần cài Microsoft Excel trên máy để chuyển đổi sang PDF.")

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
        raise ExcelComError(f"Không xuất được PDF từ Excel: {exc}")
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()
    return abs_out


def export_all_charts(path: str, tmp_dir: str) -> dict:
    """Trả về {(tên_sheet, thứ_tự_chart_0_based): đường_dẫn_png}.

    Chỉ hoạt động trên Windows (Excel COM). Trên macOS/Linux trả về {} để
    nơi gọi tự vẽ placeholder.
    """
    if sys.platform != "win32":
        return {}

    try:
        import pythoncom
        import win32com.client
    except ImportError:
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
                continue
            for chart_number in range(1, count + 1):
                try:
                    chart = chart_objects.Item(chart_number).Chart
                    safe_name = "".join(char if char.isalnum() else "_" for char in str(sheet.Name))
                    png_path = os.path.join(tmp_dir, f"chart_{safe_name}_{chart_number}.png")
                    chart.Export(png_path)
                    result[(sheet.Name, chart_number - 1)] = png_path
                except Exception:
                    continue
    except Exception:
        return {}
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()
    return result
