import sys

from PySide6.QtWidgets import QApplication

# Cấu hình logging TRƯỚC MỌI THỨ KHÁC — để log được cả lỗi xảy ra lúc import
# các module còn lại (VD thiếu thư viện, lỗi cấu hình).
from .logging_config import get_logger, log_call, setup_logging

setup_logging()
logger = get_logger(__name__)

from .account import api_client
from .app.ui import APP_STYLE, SplashWindow, MainWindow

# Giữ tham chiếu cửa sổ ở cấp module để không bị thu gom rác.
_windows: dict = {}


@log_call
def _show_main():
    win = MainWindow()
    _windows["main"] = win
    _windows.pop("splash", None)
    win.show()


@log_call
def _show_login():
    splash = SplashWindow(on_activated=_show_main)
    _windows["splash"] = splash
    splash.show()


@log_call
def restart_to_login(current_window=None):
    """Đăng xuất: xóa phiên và quay lại màn hình đăng nhập."""
    api_client.logout()
    _show_login()
    if current_window is not None:
        current_window.close()
    _windows.pop("main", None)


def main():
    logger.info("Starting application — Python %s, platform=%s",
                sys.version.split()[0], sys.platform)
    app = QApplication(sys.argv)
    app.setApplicationName("DocAI")
    app.setStyleSheet(APP_STYLE)

    # Splash đăng nhập chỉ hiện khi chưa có phiên đăng nhập
    logged_in = api_client.is_logged_in()
    logger.debug("Login state at startup: logged_in=%s", logged_in)
    if logged_in:
        _show_main()
    else:
        _show_login()

    exit_code = app.exec()
    logger.info("Application exiting with code %s", exit_code)
    sys.exit(exit_code)
