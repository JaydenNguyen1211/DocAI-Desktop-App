"""DocAI server connection config (Firebase).

Two values to fill in after deploying the server:
  - FIREBASE_API_KEY: Web API Key of the Firebase project (docai-15ceb).
    Get it at: Firebase Console → Project settings → General → Web API Key.
    This is a PUBLIC key used for login, safe to embed in the app.
  - API_BASE_URL: URL of the `api` Cloud Function after `firebase deploy`.
    Example: https://api-xxxxxx-as.a.run.app
    or:      https://asia-southeast1-docai-15ceb.cloudfunctions.net/api

Can be overridden via environment variables (DOCAI_FIREBASE_API_KEY,
DOCAI_API_BASE_URL) or the "server" key in config.json — handy for building
one .exe for multiple environments.
"""
import os

from .config import load_config

from ..logging_config import get_logger, log_call

logger = get_logger(__name__)

# ── Default values — FILL IN AFTER DEPLOYING ────────────────────────────────
FIREBASE_API_KEY = "AIzaSyDCGFsPcTfue9c04IvPZWvRtb582snbrfk"
API_BASE_URL = "https://asia-southeast1-docai-15ceb.cloudfunctions.net/api"


@log_call
def _resolve(env_key: str, cfg_key: str, default: str) -> str:
    val = os.environ.get(env_key)
    if val:
        return val.strip()
    server = load_config().get("server", {})
    if isinstance(server, dict) and server.get(cfg_key):
        return str(server[cfg_key]).strip()
    return default


@log_call
def firebase_api_key() -> str:
    return _resolve("DOCAI_FIREBASE_API_KEY", "firebase_api_key", FIREBASE_API_KEY)


@log_call
def api_base_url() -> str:
    return _resolve("DOCAI_API_BASE_URL", "api_base_url", API_BASE_URL).rstrip("/")


@log_call
def is_configured() -> bool:
    return (
        firebase_api_key() != "REPLACE_WITH_FIREBASE_WEB_API_KEY"
        and bool(firebase_api_key())
        and "REPLACE" not in api_base_url()
    )
