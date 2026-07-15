import sys

from PyQt6.QtWidgets import QApplication

from . import api_client
from .ui.styles import APP_STYLE
from .ui.splash import SplashWindow
from .ui.window import MainWindow

# Giữ tham chiếu cửa sổ ở cấp module để không bị thu gom rác.
_windows: dict = {}


def _show_main():
    win = MainWindow()
    _windows["main"] = win
    _windows.pop("splash", None)
    win.show()


def _show_login():
    splash = SplashWindow(on_activated=_show_main)
    _windows["splash"] = splash
    splash.show()


def restart_to_login(current_window=None):
    """Đăng xuất: xóa phiên và quay lại màn hình đăng nhập."""
    api_client.logout()
    _show_login()
    if current_window is not None:
        current_window.close()
    _windows.pop("main", None)


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("DocAI")
    app.setStyleSheet(APP_STYLE)

    # Splash đăng nhập chỉ hiện khi chưa có phiên đăng nhập
    if api_client.is_logged_in():
        _show_main()
    else:
        _show_login()

    sys.exit(app.exec())
