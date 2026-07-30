"""Xuất ảnh PNG của biểu đồ Excel bằng COM (Excel.Application) để ghép vào bản
xem trước tự vẽ (`renderers._render_excel_sheets` — bộ vẽ PIL riêng, không tự
vẽ được chart).

Chỉ dùng đúng 1 API COM hẹp — `Chart.Export()` — KHÔNG đụng đến vị trí/kích
thước cửa sổ Application (đó chính là nguyên nhân lỗi COM cũ trong bản trước
của app, "Unable to set the Left property of the Application class" — xem
ghi chú "Excel loi.png"). Mở file READ-ONLY và luôn `Close(SaveChanges=False)`
— không để Excel tự tính lại/"sửa" rồi ghi đè XML mà code openpyxl không
lường trước.

Không có Excel cài trên máy, hoặc lỗi COM bất kỳ → trả về {} để nơi gọi tự vẽ
placeholder, không làm hỏng cả bản xem trước.
"""
import os


class ExcelComError(Exception):
    """Không xuất được PDF qua Excel COM — thông báo tiếng Việt cho người dùng."""


def export_workbook_pdf(path: str, out_path: str) -> str:
    """Xuất toàn bộ workbook ra PDF qua Excel COM (`ExportAsFixedFormat`).

    Khác với `export_all_charts()` (tiện ích cho bản xem trước, lỗi bỏ qua
    lặng lẽ), đây là bản thân thao tác CHUYỂN ĐỔI người dùng yêu cầu — lỗi
    phải báo rõ, không im lặng trả rỗng."""
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
    except Exception as exc:  # noqa: BLE001 — mọi lỗi COM đều quy về 1 thông báo
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
    """Trả về {(tên_sheet, thứ_tự_chart_trong_sheet_0_based): đường_dẫn_png}.

    Mở đúng 1 lần Excel Application cho cả workbook (không mở lại theo từng
    chart) để giảm rủi ro treo tiến trình EXCEL.EXE.
    """
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
