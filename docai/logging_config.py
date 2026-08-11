"""Centralized logging system for the DocAI Desktop App.

Every module in `docai` shares this ONE configuration so logs have a
consistent format, a proper timestamp (date-time-millisecond, user's local
time), automatic daily file rotation, a separate error file for quick
triage, and no sensitive data (passwords/tokens) leaking into the logs.

Log file location: ``%APPDATA%/DocAI/logs/`` (Windows) or ``~/DocAI/logs``
(macOS/Linux) — same directory as `config.json` (see `account/config.py`):

    docai.log        everything from DEBUG up, rotated daily (14 days kept)
    docai_error.log  WARNING/ERROR/CRITICAL only, rotated daily (30 days kept)

Usage in any module::

    from ..logging_config import get_logger, log_call   # dot count depends on depth

    logger = get_logger(__name__)

    @log_call
    def some_function(path: str) -> bool:
        logger.debug("detail worth tracking: path=%s", path)
        ...

`setup_logging()` must be called EXACTLY ONCE, at application startup
(`docai/bootstrap.py::main()`), BEFORE the QApplication is created — every
logger created elsewhere via `get_logger()` automatically inherits this
configuration (even if that module was imported before `setup_logging()` ran).

Adjust how verbose the console output is (does not affect the log files —
they always capture everything from DEBUG) via an environment variable::

    DOCAI_LOG_LEVEL=DEBUG   (default: INFO)
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

# ── Log file location ────────────────────────────────────────────────────────

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

# Logger names used for events not tied to a specific code module.
QT_LOGGER_NAME = "docai.qt"
UNCAUGHT_LOGGER_NAME = "docai.uncaught"

_configured = False


def setup_logging(console: bool = True) -> None:
    """Configure logging for the whole app. Call once at the top of `main()`.
    Safe to call multiple times (subsequent calls are no-ops)."""
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

    # Quiet down DEBUG noise from third-party libraries (WARNING/ERROR still show).
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
    """Return the standard logger for a module — use `get_logger(__name__)`."""
    return logging.getLogger(name)


# ── Catch unforeseen errors ──────────────────────────────────────────────────

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
    """Route Qt's internal logging (qDebug/qWarning/qCritical/qFatal — e.g.
    image plugin/font errors) into the same logging system instead of just
    printing to stderr."""
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


# ── Decorator that logs function/method entry & exit ────────────────────────

_SENSITIVE_KEYS = {
    "password", "pass", "pwd", "token", "id_token", "refresh_token",
    "access_token", "api_key", "apikey", "secret", "authorization",
    "data_b64",  # base64 file content — too long, no need to see it in logs
}
_ARG_MAXLEN = 300

F = TypeVar("F", bound=Callable[..., Any])


def _safe_repr(value: Any, max_len: int = _ARG_MAXLEN) -> str:
    try:
        text = repr(value)
    except Exception:
        return f"<{type(value).__name__}: error during repr()>"
    if len(text) > max_len:
        return f"{text[:max_len]}...(truncated, {len(text)} chars long)"
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
    """Decorator that logs entry/exit of a function or method: arguments
    (sensitive data masked), run time (ms), and any exception with full
    traceback — then re-raises it unchanged (never swallows an error).

    Use the plain ``@log_call``, or customize it when the result/arguments
    are too large or sensitive to log::

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
        # Positional parameter names — needed to mask sensitive data (e.g.
        # `password`) even when the function is called positionally, not
        # by keyword.
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
    """Log the run time + any error of an arbitrary block of code (not a
    whole function) — useful for a long loop/processing block inside a
    method::

        with log_block("reading the entire Excel sheet", logger):
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
