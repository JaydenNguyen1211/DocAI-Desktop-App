"""Icon SVG nội tuyến → QIcon (theo bộ icon trong thiết kế)."""
from PyQt6.QtCore import QByteArray, Qt
from PyQt6.QtGui import QIcon, QPixmap, QPainter
from PyQt6.QtSvg import QSvgRenderer

# Đường path lấy đúng từ file thiết kế (Detailed Flows).
_PAPERCLIP = ('<path d="M21.44 11.05l-9.19 9.19a6 6 0 0 1-8.49-8.49l9.19-9.19a4 4 0 0 1 '
              '5.66 5.66l-9.2 9.19a2 2 0 0 1-2.83-2.83l8.49-8.48"/>')
_ARROW_UP = ('<line x1="12" y1="19" x2="12" y2="5"/>'
             '<polyline points="5 12 12 5 19 12"/>')


def _svg(inner: str, color: str) -> str:
    return (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" '
            f'stroke="{color}" stroke-width="2" stroke-linecap="round" '
            f'stroke-linejoin="round">{inner}</svg>')


def _render(svg: str, size: int) -> QIcon:
    renderer = QSvgRenderer(QByteArray(svg.encode("utf-8")))
    # Vẽ ở 2× cho sắc nét trên màn hình HiDPI.
    pm = QPixmap(size * 2, size * 2)
    pm.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pm)
    renderer.render(painter)
    painter.end()
    pm.setDevicePixelRatio(2.0)
    return QIcon(pm)


def paperclip_icon(color: str = "#6E6A63", size: int = 17) -> QIcon:
    return _render(_svg(_PAPERCLIP, color), size)


def diamond_icon(color: str, size: int = 11) -> QIcon:
    svg = (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">'
           f'<path d="M12 3 L21 12 L12 21 L3 12 Z" fill="{color}"/></svg>')
    return _render(svg, size)


def send_icon(color: str = "#FFFFFF", size: int = 15) -> QIcon:
    return _render(_svg(_ARROW_UP, color), size)
