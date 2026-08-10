"""Client gọi API server DocAI.

- Đăng ký / đăng nhập qua Firebase Authentication (Identity Toolkit REST).
- Gọi chat AI và lấy thông tin tài khoản qua Cloud Function `api` (server proxy).
Token được lưu vào config.json để giữ đăng nhập giữa các phiên.
"""
import time

import requests

from .config import load_config, save_config
from .server_config import firebase_api_key, api_base_url

from ..logging_config import get_logger, log_call

logger = get_logger(__name__)

_IDENTITY = "https://identitytoolkit.googleapis.com/v1/accounts"
_TOKEN_URL = "https://securetoken.googleapis.com/v1/token"
_TIMEOUT = 300  # AI có thể mất vài chục giây

# Ánh xạ lỗi Firebase → thông báo tiếng Việt.
_ERR_VI = {
    "EMAIL_EXISTS": "Email này đã được đăng ký.",
    "INVALID_EMAIL": "Email không hợp lệ.",
    "EMAIL_NOT_FOUND": "Email chưa được đăng ký.",
    "INVALID_PASSWORD": "Sai mật khẩu.",
    "INVALID_LOGIN_CREDENTIALS": "Email hoặc mật khẩu không đúng.",
    "MISSING_PASSWORD": "Vui lòng nhập mật khẩu.",
    "USER_DISABLED": "Tài khoản đã bị khóa.",
    "TOO_MANY_ATTEMPTS_TRY_LATER":
        "Đăng nhập sai quá nhiều lần, vui lòng thử lại sau.",
}


class ApiError(Exception):
    """Lỗi có thông báo tiếng Việt để hiển thị cho người dùng."""

    @log_call
    def __init__(self, message: str, code: str = ""):
        super().__init__(message)
        self.message = message
        self.code = code


@log_call
def _friendly(code: str) -> str:
    base = code.split(":")[0].strip()  # WEAK_PASSWORD : ... → WEAK_PASSWORD
    if base == "WEAK_PASSWORD":
        return "Mật khẩu quá yếu (cần tối thiểu 6 ký tự)."
    return _ERR_VI.get(base, f"Lỗi xác thực: {code}" if code else "Lỗi không rõ.")


# ── Lưu / đọc phiên đăng nhập ──────────────────────────────────────────────

@log_call(log_args=False)  # `data` chứa idToken/refreshToken — không log
def _store_session(data: dict):
    """Lưu phiên từ cả 2 nguồn: Identity Toolkit (idToken/refreshToken/expiresIn)
    và securetoken refresh (id_token/refresh_token/expires_in)."""
    cfg = load_config()
    old = cfg.get("auth", {})
    id_token = data.get("idToken") or data.get("id_token")
    refresh_token = data.get("refreshToken") or data.get("refresh_token")
    if not id_token or not refresh_token:
        raise ApiError("Phiên đăng nhập không hợp lệ, vui lòng đăng nhập lại.",
                       "BAD_SESSION")
    expires_in = data.get("expiresIn") or data.get("expires_in") or "3600"
    cfg["auth"] = {
        "id_token": id_token,
        "refresh_token": refresh_token,
        "uid": data.get("localId") or data.get("user_id") or old.get("uid", ""),
        "email": data.get("email", old.get("email", "")),
        # expiresIn (giây) → mốc hết hạn tuyệt đối
        "expires_at": time.time() + int(expires_in) - 60,
    }
    save_config(cfg)


@log_call(log_result=False)  # kết quả chứa auth token — không log
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

@log_call(log_args=False, log_result=False)  # payload/kết quả chứa password/token
def _identity(endpoint: str, payload: dict) -> dict:
    key = firebase_api_key()
    try:
        response = requests.post(
            f"{_IDENTITY}:{endpoint}?key={key}", json=payload, timeout=30,
        )
    except requests.RequestException:
        raise ApiError("Không kết nối được máy chủ. Kiểm tra mạng.")
    body = response.json() if response.content else {}
    if response.status_code >= 400:
        code = (body.get("error", {}) or {}).get("message", "")
        raise ApiError(_friendly(code), code)
    return body


@log_call(log_result=False)  # kết quả chứa idToken/refreshToken
def signup(email: str, password: str) -> dict:
    data = _identity("signUp", {
        "email": email, "password": password, "returnSecureToken": True,
    })
    _store_session(data)
    return data


@log_call(log_result=False)  # kết quả chứa idToken/refreshToken
def login(email: str, password: str) -> dict:
    data = _identity("signInWithPassword", {
        "email": email, "password": password, "returnSecureToken": True,
    })
    _store_session(data)
    return data


@log_call(log_result=False)  # trả về id_token thô
def _refresh() -> str:
    sess = _session()
    token = sess.get("refresh_token")
    if not token:
        raise ApiError("Chưa đăng nhập.", "NO_SESSION")
    try:
        response = requests.post(
            f"{_TOKEN_URL}?key={firebase_api_key()}",
            data={"grant_type": "refresh_token", "refresh_token": token},
            timeout=30,
        )
    except requests.RequestException:
        raise ApiError("Không kết nối được máy chủ. Kiểm tra mạng.")
    if response.status_code >= 400:
        logout()
        raise ApiError("Phiên đăng nhập hết hạn, vui lòng đăng nhập lại.",
                       "REFRESH_FAILED")
    _store_session(response.json())
    return _session()["id_token"]


@log_call(log_result=False)  # trả về id_token thô
def _valid_token() -> str:
    sess = _session()
    if not sess.get("id_token"):
        return _refresh()
    if time.time() >= sess.get("expires_at", 0):
        return _refresh()
    return sess["id_token"]


# ── Gọi server (Cloud Function `api`) ──────────────────────────────────────

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
        raise ApiError("Không kết nối được máy chủ. Kiểm tra mạng.")

    if response.status_code == 401 and retry_auth:
        logger.debug("401 from %s %s — refreshing token and retrying once", method, path)
        _refresh()
        return _call(method, path, json_body, retry_auth=False)

    body = response.json() if response.content else {}
    if response.status_code == 402:
        raise ApiError("Đã hết lượt gọi AI trong tháng. Nâng cấp gói để tiếp tục.",
                       "quota_exceeded")
    if response.status_code >= 400 or not body.get("ok", False):
        raise ApiError(f"Máy chủ báo lỗi ({response.status_code}).",
                       body.get("error", "server_error"))
    return body


@log_call
def me() -> dict:
    """Trả về {plan, quota_used, quota_limit, quota_remaining, business, email}."""
    return _call("GET", "/me")


@log_call
def update_business(business: dict) -> dict:
    return _call("POST", "/business", {"business": business})


@log_call
def get_rates() -> dict:
    """Tham số tính lương/BHXH/thuế hiện hành theo server (`{ok, source,
    rates}`) — dùng bởi `modules.business.payroll`. Không trừ quota."""
    return _call("GET", "/rates")


@log_call
def chat(message: str, file_name: str = "", file_type: str = "",
         business: dict | None = None, history: list | None = None,
         attachment: dict | None = None) -> dict:
    """Gọi AI qua server. Trả về {text, plan, quota_remaining}.

    `attachment` (tùy chọn): {"kind": "pdf"|"image", "media_type": ...,
    "data_b64": ...} — nội dung file PDF/ảnh đang mở, để Claude đọc trực tiếp
    (PDF: cả bản có chữ lẫn bản quét ảnh; ảnh: qua Claude Vision).
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
    """Nhờ AI đọc PDF/ảnh (hóa đơn, hợp đồng, bảng biểu…) và trả về nội dung
    dạng bảng theo quy ước "--- Sheet: X ---" + CSV — cùng định dạng luồng
    "tạo tài liệu mới bằng chat" đã dùng, nên chỉ cần gọi thẳng
    `document.create_excel_from_text()` với text trả về, không cần parse riêng.
    Trả về {ok, text, plan, quota_remaining}.
    """
    return _call("POST", "/extract_table", {
        "attachment": attachment,
        "message": message or "",
    })


@log_call
def extract_text(attachment: dict, message: str = "") -> dict:
    """Nhờ AI đọc (OCR) toàn bộ chữ trong ảnh chụp tài liệu, trả về văn bản
    THUẦN giữ đúng cách ngắt đoạn gốc — khác `extract_table` (chỉ dành cho dữ
    liệu dạng bảng). Dùng cho "Ảnh → Word chỉnh sửa được" và PDF có lớp text ẩn.
    Trả về {ok, text, plan, quota_remaining}.
    """
    return _call("POST", "/extract_text", {
        "attachment": attachment,
        "message": message or "",
    })


@log_call
def edit_file(message: str, file_name: str, file_type: str,
              outline: str = "", history: list | None = None,
              attachment: dict | None = None) -> dict:
    """Hỏi AI danh sách lệnh sửa file (word/excel/ppt/image).

    `outline` là cấu trúc tài liệu (đoạn/trang/ô/slide) để AI nhắm đúng vị trí.
    `attachment` (tùy chọn, dùng cho ảnh): nội dung ảnh để Claude vẫn trả lời
    được câu hỏi không phải lệnh sửa, xem `chat()`.
    Trả về {edits, reply, plan, quota_remaining} — app tự áp lệnh lên file;
    `edits` rỗng nghĩa là câu chat không phải yêu cầu sửa.
    """
    return _call("POST", "/edit", {
        "message": message,
        "file_name": file_name,
        "file_type": file_type,
        "outline": outline,
        "history": history or [],
        "attachment": attachment,
    })
