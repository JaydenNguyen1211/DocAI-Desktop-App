"""Client for calling the DocAI API server.

- Sign up / log in via Firebase Authentication (Identity Toolkit REST).
- Call the AI chat and fetch account info via the `api` Cloud Function (server proxy).
Tokens are stored in config.json to keep the user logged in across sessions.
"""
import time

import requests

from .config import load_config, save_config
from .server_config import firebase_api_key, api_base_url

from ..logging_config import get_logger, log_call
from ..strings import Account as S

logger = get_logger(__name__)

_IDENTITY = "https://identitytoolkit.googleapis.com/v1/accounts"
_TOKEN_URL = "https://securetoken.googleapis.com/v1/token"
_TIMEOUT = 300  # AI calls can take tens of seconds


class ApiError(Exception):
    """Error carrying a Vietnamese message to show the user."""

    @log_call
    def __init__(self, message: str, code: str = ""):
        super().__init__(message)
        self.message = message
        self.code = code


@log_call
def _friendly(code: str) -> str:
    base = code.split(":")[0].strip()  # WEAK_PASSWORD : ... → WEAK_PASSWORD
    if base == "WEAK_PASSWORD":
        return S.WEAK_PASSWORD
    return S.ERROR_MAP.get(base, S.AUTH_ERROR.format(code=code) if code else S.UNKNOWN_ERROR)


# ── Save / read the login session ───────────────────────────────────────────

@log_call(log_args=False)  # `data` contains idToken/refreshToken — don't log
def _store_session(data: dict):
    """Store the session from either source: Identity Toolkit
    (idToken/refreshToken/expiresIn) or securetoken refresh
    (id_token/refresh_token/expires_in)."""
    cfg = load_config()
    old = cfg.get("auth", {})
    id_token = data.get("idToken") or data.get("id_token")
    refresh_token = data.get("refreshToken") or data.get("refresh_token")
    if not id_token or not refresh_token:
        raise ApiError(S.INVALID_SESSION, "BAD_SESSION")
    expires_in = data.get("expiresIn") or data.get("expires_in") or "3600"
    cfg["auth"] = {
        "id_token": id_token,
        "refresh_token": refresh_token,
        "uid": data.get("localId") or data.get("user_id") or old.get("uid", ""),
        "email": data.get("email", old.get("email", "")),
        # expiresIn (seconds) → absolute expiry timestamp
        "expires_at": time.time() + int(expires_in) - 60,
    }
    save_config(cfg)


@log_call(log_result=False)  # result contains the auth token — don't log
def _session() -> dict:
    return load_config().get("auth", {})


@log_call
def is_logged_in() -> bool:
    return bool(_session().get("refresh_token"))


@log_call
def current_email() -> str:
    return _session().get("email", "")


@log_call
def logout():
    cfg = load_config()
    cfg.pop("auth", None)
    save_config(cfg)


# ── Firebase Identity Toolkit ──────────────────────────────────────────────

@log_call(log_args=False, log_result=False)  # payload/result contains password/token
def _identity(endpoint: str, payload: dict) -> dict:
    key = firebase_api_key()
    try:
        response = requests.post(
            f"{_IDENTITY}:{endpoint}?key={key}", json=payload, timeout=30,
        )
    except requests.RequestException:
        raise ApiError(S.NO_CONNECTION)
    body = response.json() if response.content else {}
    if response.status_code >= 400:
        code = (body.get("error", {}) or {}).get("message", "")
        raise ApiError(_friendly(code), code)
    return body


@log_call(log_result=False)  # result contains idToken/refreshToken
def signup(email: str, password: str) -> dict:
    data = _identity("signUp", {
        "email": email, "password": password, "returnSecureToken": True,
    })
    _store_session(data)
    return data


@log_call(log_result=False)  # result contains idToken/refreshToken
def login(email: str, password: str) -> dict:
    data = _identity("signInWithPassword", {
        "email": email, "password": password, "returnSecureToken": True,
    })
    _store_session(data)
    return data


@log_call(log_result=False)  # returns the raw id_token
def _refresh() -> str:
    sess = _session()
    token = sess.get("refresh_token")
    if not token:
        raise ApiError(S.NOT_LOGGED_IN, "NO_SESSION")
    try:
        response = requests.post(
            f"{_TOKEN_URL}?key={firebase_api_key()}",
            data={"grant_type": "refresh_token", "refresh_token": token},
            timeout=30,
        )
    except requests.RequestException:
        raise ApiError(S.NO_CONNECTION)
    if response.status_code >= 400:
        logout()
        raise ApiError(S.SESSION_EXPIRED, "REFRESH_FAILED")
    _store_session(response.json())
    return _session()["id_token"]


@log_call(log_result=False)  # returns the raw id_token
def _valid_token() -> str:
    sess = _session()
    if not sess.get("id_token"):
        return _refresh()
    if time.time() >= sess.get("expires_at", 0):
        return _refresh()
    return sess["id_token"]


# ── Call the server (Cloud Function `api`) ──────────────────────────────────

@log_call
def _call(method: str, path: str, json_body: dict | None = None,
          retry_auth: bool = True) -> dict:
    token = _valid_token()
    url = f"{api_base_url()}{path}"
    try:
        response = requests.request(
            method, url, json=json_body, timeout=_TIMEOUT,
            headers={"Authorization": f"Bearer {token}"},
        )
    except requests.RequestException:
        raise ApiError(S.NO_CONNECTION)

    if response.status_code == 401 and retry_auth:
        logger.debug("401 from %s %s — refreshing token and retrying once", method, path)
        _refresh()
        return _call(method, path, json_body, retry_auth=False)

    body = response.json() if response.content else {}
    if response.status_code == 402:
        raise ApiError(S.QUOTA_EXCEEDED, "quota_exceeded")
    if response.status_code >= 400 or not body.get("ok", False):
        raise ApiError(S.SERVER_ERROR.format(status=response.status_code),
                       body.get("error", "server_error"))
    return body


@log_call
def me() -> dict:
    """Returns {plan, quota_used, quota_limit, quota_remaining, business, email}."""
    return _call("GET", "/me")


@log_call
def update_business(business: dict) -> dict:
    return _call("POST", "/business", {"business": business})


@log_call
def get_rates() -> dict:
    """Current payroll/social-insurance/tax parameters from the server
    (`{ok, source, rates}`) — used by `modules.business.payroll`. Does not
    consume quota."""
    return _call("GET", "/rates")


@log_call
def chat(message: str, file_name: str = "", file_type: str = "",
         business: dict | None = None, history: list | None = None,
         attachment: dict | None = None) -> dict:
    """Call the AI via the server. Returns {text, plan, quota_remaining}.

    `attachment` (optional): {"kind": "pdf"|"image", "media_type": ...,
    "data_b64": ...} — content of the currently open PDF/image, so Claude
    can read it directly (PDF: both text and scanned-image; image: via
    Claude Vision).
    """
    return _call("POST", "/chat", {
        "message": message,
        "file_name": file_name or "",
        "file_type": file_type or "",
        "business": business or {},
        "history": history or [],
        "attachment": attachment,
    })


@log_call
def extract_table(attachment: dict, message: str = "") -> dict:
    """Ask the AI to read a PDF/image (invoice, contract, table…) and return
    the content as a table, using the "--- Sheet: X ---" + CSV convention —
    the same format the "create new document via chat" flow already uses,
    so the returned text can be passed straight to
    `document.create_excel_from_text()` with no separate parsing needed.
    Returns {ok, text, plan, quota_remaining}.
    """
    return _call("POST", "/extract_table", {
        "attachment": attachment,
        "message": message or "",
    })


@log_call
def extract_text(attachment: dict, message: str = "") -> dict:
    """Ask the AI to OCR all the text out of a scanned document image,
    returning PLAIN text that preserves the original paragraph breaks —
    unlike `extract_table` (table data only). Used for "Image → editable
    Word" and PDFs with a hidden text layer.
    Returns {ok, text, plan, quota_remaining}.
    """
    return _call("POST", "/extract_text", {
        "attachment": attachment,
        "message": message or "",
    })


@log_call
def edit_file(message: str, file_name: str, file_type: str,
              outline: str = "", history: list | None = None,
              attachment: dict | None = None) -> dict:
    """Ask the AI for a list of edit commands for a file (word/excel/ppt/image).

    `outline` is the document structure (paragraph/page/cell/slide) so the AI
    can target the right location.
    `attachment` (optional, used for images): the image content so Claude
    can still answer questions that aren't edit commands, see `chat()`.
    Returns {edits, reply, plan, quota_remaining} — the app applies the
    commands to the file itself; an empty `edits` means the chat message
    wasn't an edit request.
    """
    return _call("POST", "/edit", {
        "message": message,
        "file_name": file_name,
        "file_type": file_type,
        "outline": outline,
        "history": history or [],
        "attachment": attachment,
    })
