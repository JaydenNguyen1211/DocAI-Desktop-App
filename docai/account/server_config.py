"""Cấu hình kết nối server DocAI (Firebase).

Ba giá trị cần điền sau khi deploy server:
  - FIREBASE_API_KEY: Web API Key của project Firebase (docai-15ceb).
    Lấy tại: Firebase Console → Project settings → General → Web API Key.
    Đây là khóa CÔNG KHAI dùng cho đăng nhập, an toàn khi nhúng vào app.
  - API_BASE_URL: URL của Cloud Function `api` sau khi `firebase deploy`.
    Ví dụ: https://api-xxxxxx-as.a.run.app
    hoặc:  https://asia-southeast1-docai-15ceb.cloudfunctions.net/api

Có thể override bằng biến môi trường (DOCAI_FIREBASE_API_KEY, DOCAI_API_BASE_URL)
hoặc bằng khóa "server" trong config.json — tiện khi build .exe cho nhiều môi trường.
"""
import os

from .config import load_config

# ── Giá trị mặc định — ĐIỀN SAU KHI DEPLOY ─────────────────────────────────
FIREBASE_API_KEY = "AIzaSyDCGFsPcTfue9c04IvPZWvRtb582snbrfk"
API_BASE_URL = "https://asia-southeast1-docai-15ceb.cloudfunctions.net/api"


def _resolve(env_key: str, cfg_key: str, default: str) -> str:
    val = os.environ.get(env_key)
    if val:
        return val.strip()
    server = load_config().get("server", {})
    if isinstance(server, dict) and server.get(cfg_key):
        return str(server[cfg_key]).strip()
    return default


def firebase_api_key() -> str:
    return _resolve("DOCAI_FIREBASE_API_KEY", "firebase_api_key", FIREBASE_API_KEY)


def api_base_url() -> str:
    return _resolve("DOCAI_API_BASE_URL", "api_base_url", API_BASE_URL).rstrip("/")


def is_configured() -> bool:
    return (
        firebase_api_key() != "REPLACE_WITH_FIREBASE_WEB_API_KEY"
        and bool(firebase_api_key())
        and "REPLACE" not in api_base_url()
    )
