import json
import os
from pathlib import Path

CONFIG_PATH = Path(os.environ.get("APPDATA", Path.home())) / "DocAI" / "config.json"


def load_config() -> dict:
    try:
        if CONFIG_PATH.exists():
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}


def save_config(data: dict):
    try:
        CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
        CONFIG_PATH.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except Exception:
        pass
