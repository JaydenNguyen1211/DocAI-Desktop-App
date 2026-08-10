"""Hệ thống log tập trung cho DocAI Desktop App.

Mọi module trong `docai` dùng CHUNG cấu hình này để log có định dạng thống
nhất, có timestamp chuẩn (ngày-giờ-mili giây, giờ local máy người dùng),
tự động xoay vòng file theo ngày, tách riêng file lỗi để soi nhanh, và
không để lọt dữ liệu nhạy cảm (mật khẩu/token) ra log.

Vị trí file log: ``%APPDATA%/DocAI/logs/`` (Windows) hoặc ``~/DocAI/logs``
(macOS/Linux) — cùng thư mục với `config.json` (xem `account/config.py`):

    docai.log        toàn bộ log từ DEBUG trở lên, xoay theo ngày (giữ 14 ngày)
    docai_error.log  chỉ WARNING/ERROR/CRITICAL, xoay theo ngày (giữ 30 ngày)

Cách dùng trong 1 module bất kỳ::

    from ..logging_config import get_logger, log_call   # số dấu chấm tùy độ sâu

    logger = get_logger(__name__)

    @log_call
    def some_function(path: str) -> bool:
        logger.debug("detail worth tracking: path=%s", path)
        ...

`setup_logging()` chỉ gọi 1 LẦN DUY NHẤT ở điểm khởi động ứng dụng
(`docai/bootstrap.py::main()`), TRƯỚC khi tạo QApplication — mọi logger tạo
bằng `get_logger()` ở nơi khác tự động thừa hưởng cấu hình này (kể cả khi
module đó được import trước khi `setup_logging()` chạy).

Điều chỉnh độ chi tiết hiện trên console (không ảnh hưởng file log — file
luôn ghi đầy đủ từ DEBUG) bằng biến môi trường::

    DOCAI_LOG_LEVEL=DEBUG   (mặc định: INFO)
"""
from __future__ import annotations

import functools
import inspect
import logging
import logging.handlers
import os
import sys
import time
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Callable, Optional, TypeVar

# ── Vị trí file log ──────────────────────────────────────────────────────────

def _log_dir() -> Path:
    base = os.environ.get("APPDATA") or str(Path.home())
    return Path(base) / "DocAI" / "logs"


LOG_DIR = _log_dir()
MAIN_LOG_FILE = LOG_DIR / "docai.log"
ERROR_LOG_FILE = LOG_DIR / "docai_error.log"

_DATEFMT = "%Y-%m-%d %H:%M:%S"
_FILE_FMT = (
    "%(asctime)s.%(msecs)03d | %(levelname)-8s | %(threadName)-15s | "
    "%(name)s:%(funcName)s:%(lineno)d | %(message)s"
)
_CONSOLE_FMT = "%(asctime)s.%(msecs)03d | %(levelname)-8s | %(name)s | %(message)s"

# Tên logger dùng cho các sự kiện không gắn với 1 module code cụ thể.
QT_LOGGER_NAME = "docai.qt"
UNCAUGHT_LOGGER_NAME = "docai.uncaught"

_configured = False


def setup_logging(console: bool = True) -> None:
    """Cấu hình logging cho toàn bộ app. Gọi 1 lần ở đầu `main()`. An toàn khi
    gọi nhiều lần (lần sau bị bỏ qua)."""
    global _configured
    if _configured:
        return
    _configured = True

    root = logging.getLogger()
    root.setLevel(logging.DEBUG)

    file_fmt = logging.Formatter(fmt=_FILE_FMT, datefmt=_DATEFMT)

    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)

        main_handler = logging.handlers.TimedRotatingFileHandler(
            str(MAIN_LOG_FILE), when="midnight", backupCount=14,
            encoding="utf-8", delay=True,
        )
        main_handler.suffix = "%Y-%m-%d"
        main_handler.setLevel(logging.DEBUG)
        main_handler.setFormatter(file_fmt)
        root.addHandler(main_handler)

        error_handler = logging.handlers.TimedRotatingFileHandler(
            str(ERROR_LOG_FILE), when="midnight", backupCount=30,
            encoding="utf-8", delay=True,
        )
        error_handler.suffix = "%Y-%m-%d"
        error_handler.setLevel(logging.WARNING)
        error_handler.setFormatter(file_fmt)
        root.addHandler(error_handler)
    except OSError:
        logging.getLogger(__name__).warning(
            "Could not create log directory %s — logging to console only.",
            LOG_DIR, exc_info=True,
        )

    if console:
        stream = sys.stdout or sys.stderr
        if stream is not None and hasattr(stream, "reconfigure"):
            # Windows console defaults to a legacy codepage (e.g. cp1252) —
            # log messages contain Vietnamese comments/paths, so force UTF-8
            # to avoid mangled output. Best-effort: never fail startup on this.
            try:
                stream.reconfigure(encoding="utf-8", errors="backslashreplace")
            except (ValueError, OSError):
                pass
        if stream is not None:
            level_name = os.environ.get("DOCAI_LOG_LEVEL", "INFO").upper()
            console_level = getattr(logging, level_name, logging.INFO)
            console_handler = logging.StreamHandler(stream)
            console_handler.setLevel(console_level)
            console_handler.setFormatter(
                logging.Formatter(fmt=_CONSOLE_FMT, datefmt=_DATEFMT))
            root.addHandler(console_handler)

    # Giảm nhiễu log DEBUG từ thư viện bên thứ 3 (vẫn thấy WARNING/ERROR).
    for noisy in ("urllib3", "PIL", "fontTools", "comtypes", "charset_normalizer"):
        logging.getLogger(noisy).setLevel(logging.WARNING)

    install_excepthook()
    try:
        install_qt_message_handler()
    except Exception:
        logging.getLogger(__name__).debug(
            "Could not install Qt message handler (skipped).", exc_info=True)

    logging.getLogger(__name__).info(
        "=== DocAI starting — logs at: %s ===", LOG_DIR)


def get_logger(name: str) -> logging.Logger:
    """Trả về logger chuẩn cho 1 module — dùng `get_logger(__name__)`."""
    return logging.getLogger(name)


# ── Bắt lỗi không lường trước ────────────────────────────────────────────────

def _excepthook(exc_type, exc_value, exc_tb):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_tb)
        return
    logging.getLogger(UNCAUGHT_LOGGER_NAME).critical(
        "Uncaught exception — application state may be unstable",
        exc_info=(exc_type, exc_value, exc_tb),
    )
    sys.__excepthook__(exc_type, exc_value, exc_tb)


def install_excepthook() -> None:
    sys.excepthook = _excepthook


def install_qt_message_handler() -> None:
    """Chuyển log nội bộ của Qt (qDebug/qWarning/qCritical/qFatal — VD lỗi
    plugin ảnh, font, v.v.) vào cùng hệ thống log thay vì chỉ in ra stderr."""
    from PySide6.QtCore import QtMsgType, qInstallMessageHandler

    qt_logger = logging.getLogger(QT_LOGGER_NAME)
    _LEVEL_MAP = {
        QtMsgType.QtDebugMsg: logging.DEBUG,
        QtMsgType.QtInfoMsg: logging.INFO,
        QtMsgType.QtWarningMsg: logging.WARNING,
        QtMsgType.QtCriticalMsg: logging.ERROR,
        QtMsgType.QtFatalMsg: logging.CRITICAL,
    }

    def _handler(msg_type, _context, message):
        qt_logger.log(_LEVEL_MAP.get(msg_type, logging.INFO), "%s", message)

    qInstallMessageHandler(_handler)


# ── Decorator log vào/ra hàm & method ───────────────────────────────────────

_SENSITIVE_KEYS = {
    "password", "pass", "pwd", "token", "id_token", "refresh_token",
    "access_token", "api_key", "apikey", "secret", "authorization",
    "data_b64",  # nội dung file base64 — quá dài, không cần thấy trong log
}
_ARG_MAXLEN = 300

F = TypeVar("F", bound=Callable[..., Any])


def _safe_repr(value: Any, max_len: int = _ARG_MAXLEN) -> str:
    try:
        text = repr(value)
    except Exception:
        return f"<{type(value).__name__}: lỗi khi repr()>"
    if len(text) > max_len:
        return f"{text[:max_len]}...(cắt bớt, dài {len(text)} ký tự)"
    return text


def _mask(key: str, value: Any) -> str:
    if key.lower() in _SENSITIVE_KEYS:
        return "***"
    return _safe_repr(value)


def log_call(
    func: Optional[F] = None,
    *,
    level: int = logging.DEBUG,
    log_args: bool = True,
    log_result: bool = True,
) -> Any:
    """Decorator ghi log khi vào/ra 1 hàm hoặc method: tham số (đã che dữ
    liệu nhạy cảm), thời gian chạy (ms), và exception kèm traceback đầy đủ
    nếu có — rồi re-raise nguyên vẹn (không nuốt lỗi).

    Dùng mặc định ``@log_call``, hoặc tùy biến khi log kết quả/tham số quá
    lớn hay nhạy cảm::

        @log_call(log_result=False)
        def read_file_bytes(path: str) -> bytes: ...
    """

    def decorator(fn: F) -> F:
        target_logger = logging.getLogger(fn.__module__)
        qualname = fn.__qualname__

        try:
            params = list(inspect.signature(fn).parameters)
        except (TypeError, ValueError):
            params = []
        is_method = bool(params) and params[0] in ("self", "cls")
        # Tên tham số theo vị trí — để che dữ liệu nhạy cảm (VD `password`)
        # kể cả khi gọi hàm theo kiểu positional thay vì keyword.
        pos_names = params[1:] if is_method else params

        @functools.wraps(fn)
        def wrapper(*args, **kwargs):
            if log_args:
                pos = args[1:] if is_method else args
                parts = [
                    _mask(pos_names[i], v) if i < len(pos_names) else _safe_repr(v)
                    for i, v in enumerate(pos)
                ]
                parts += [f"{k}={_mask(k, v)}" for k, v in kwargs.items()]
                params_str = ", ".join(parts)
            else:
                params_str = "..."

            target_logger.log(level, "→ %s(%s)", qualname, params_str)
            start = time.perf_counter()
            try:
                result = fn(*args, **kwargs)
            except Exception:
                elapsed = (time.perf_counter() - start) * 1000
                target_logger.exception(
                    "✗ %s FAILED after %.1fms", qualname, elapsed)
                raise
            elapsed = (time.perf_counter() - start) * 1000
            if log_result:
                target_logger.log(
                    level, "← %s done in %.1fms → %s",
                    qualname, elapsed, _safe_repr(result))
            else:
                target_logger.log(
                    level, "← %s done in %.1fms", qualname, elapsed)
            return result

        return wrapper  # type: ignore[return-value]

    if func is not None:
        return decorator(func)
    return decorator


@contextmanager
def log_block(name: str, logger: Optional[logging.Logger] = None,
              level: int = logging.DEBUG):
    """Log thời gian chạy + lỗi của 1 đoạn code bất kỳ (không phải cả hàm) —
    dùng cho vòng lặp/khối xử lý dài bên trong 1 method::

        with log_block("đọc toàn bộ sheet Excel", logger):
            for row in sheet.iter_rows():
                ...
    """
    log = logger or logging.getLogger("docai")
    start = time.perf_counter()
    log.log(level, "▶ start: %s", name)
    try:
        yield
    except Exception:
        elapsed = (time.perf_counter() - start) * 1000
        log.exception("✗ error in '%s' after %.1fms", name, elapsed)
        raise
    else:
        elapsed = (time.perf_counter() - start) * 1000
        log.log(level, "■ done: %s (%.1fms)", name, elapsed)
