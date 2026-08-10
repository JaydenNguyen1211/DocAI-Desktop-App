import json
import os
from pathlib import Path

from ..logging_config import get_logger, log_call

logger = get_logger(__name__)

CONFIG_PATH = Path(os.environ.get("APPDATA", Path.home())) / "DocAI" / "config.json"


@log_call(log_result=False)  # kết quả chứa auth token — không log ra file
def load_config() -> dict:
    try:
        if CONFIG_PATH.exists():
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception:
        logger.exception("Failed to read/parse config file at %s", CONFIG_PATH)
    return {}


@log_call(log_args=False)  # `data` chứa auth token — không log ra file
def save_config(data: dict):
    try:
        CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
        CONFIG_PATH.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except Exception:
        logger.exception("Failed to write config file at %s", CONFIG_PATH)
