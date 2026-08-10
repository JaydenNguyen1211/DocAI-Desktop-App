"""Tạo tài liệu mới bằng chat.

`detect_create_intent()` quyết định có phải yêu cầu TẠO file mới hay không
(khác với sửa file đang mở, vốn đã được lọc trước ở app.ui.main_window) và
suy đoán định dạng nếu câu lệnh có nhắc tới. Việc soạn nội dung thật do server
(Claude, qua /chat có sẵn) đảm nhiệm — hàm này chỉ điều hướng UI.
"""
import re
from typing import Optional

from ..common.filenames import FORBIDDEN_FILENAME_CHARS

from ...logging_config import get_logger, log_call

logger = get_logger(__name__)

_CREATE_VERBS = ("tạo", "soạn", "lập", "làm", "viết", "dựng", "thiết kế", "vẽ")
_GENERIC_NOUNS = ("file", "tài liệu", "giấy tờ", "văn bản")
_PPT_WORDS = ("slide", "powerpoint", "trình bày", "thuyết trình", "ppt", ".pptx")
_EXCEL_WORDS = (
    "excel", "bảng tính", "bảng lương", "bảng kê", "bảng theo dõi",
    "spreadsheet", "xls", ".xlsx", "công nợ", "chấm công",
)
_WORD_WORDS = (
    "công văn", "tờ trình", "quyết định", "thông báo", "hợp đồng",
    "đơn xin", "thư mời", "biên bản", ".docx",
)
_DOMAIN_HINTS = (
    "hóa đơn", "hoá đơn", "lương", "chấm công", "thuế", "gtgt", "tncn",
    "hợp đồng", "công văn", "tờ trình", "quyết định", "thông báo",
    "giới thiệu", "trình bày", "slide", "kế hoạch", "dự án", "doanh thu",
    "tiến độ", "sự cố", "công nợ", "nhân sự", "báo giá", "biên bản",
)

_STOPWORDS = {
    "tạo", "soạn", "lập", "làm", "viết", "dựng", "thiết", "kế", "cho",
    "mình", "giúp", "tôi", "một", "cái", "bản", "file", "tài", "liệu",
    "trang", "slide", "và", "về", "với", "là", "của", "trình", "bày", "vẽ",
}


@log_call
def detect_create_intent(text: str, has_open_file: bool = False) -> Optional[str]:
    """Trả về "word"/"excel"/"ppt" (định dạng suy đoán được), "needs_type"
    (chắc chắn là yêu cầu tạo mới nhưng chưa rõ định dạng), "ambiguous"
    (câu lệnh quá mơ hồ — cần AI hỏi lại, xem EC6), hoặc None (không phải
    yêu cầu tạo tài liệu mới)."""
    low = text.lower()

    if not any(verb in low for verb in _CREATE_VERBS):
        return None
    if has_open_file and "mới" not in low:
        return None

    if any(word in low for word in _PPT_WORDS):
        return "ppt"
    if any(word in low for word in _EXCEL_WORDS):
        return "excel"
    if any(word in low for word in _WORD_WORDS):
        return "word"

    if _is_ambiguous(low):
        return "ambiguous"
    return "needs_type"


@log_call
def _is_ambiguous(low: str) -> bool:
    has_generic_noun = any(noun in low for noun in _GENERIC_NOUNS)
    has_domain = any(domain_word in low for domain_word in _DOMAIN_HINTS)
    word_count = len(low.split())
    return has_generic_noun and not has_domain and word_count <= 8


@log_call
def build_creation_prompt(file_type: str, user_text: str, page_count: int,
                          business: dict | None = None) -> str:
    """Câu lệnh gửi lên server (/chat) — yêu cầu AI trả về nội dung kèm mốc
    tách trang/sheet để `modules.common.creators.parse_sections()` /
    `save_excel()` dựng file thật."""
    company = (business or {}).get("company_name") or ""
    biz_line = f"Thông tin doanh nghiệp (dùng nếu phù hợp): {company}.\n" if company else ""

    if file_type == "excel":
        return (
            f"{user_text}\n\n{biz_line}"
            "Hãy soạn nội dung dạng bảng tính Excel. Với MỖI sheet, bắt đầu "
            "bằng một dòng riêng đúng định dạng:\n"
            "--- Sheet: <tên sheet> ---\n"
            "sau đó là các dòng dữ liệu, mỗi dòng là 1 hàng, các ô cách nhau "
            "bằng dấu phẩy (CSV). Không thêm giải thích ngoài các dòng trên."
        )

    noun = "slide" if file_type == "ppt" else "trang"
    return (
        f"{user_text}\n\n{biz_line}"
        f"Hãy soạn nội dung khoảng {page_count} {noun}. Với MỖI {noun}, bắt "
        "đầu bằng một dòng riêng đúng định dạng:\n"
        f"### Trang <số thứ tự>: <tiêu đề {noun}>\n"
        f"sau đó là nội dung {noun} đó (vài dòng ngắn gọn, súc tích). "
        "Không thêm giải thích ngoài các dòng trên."
    )


@log_call
def guess_page_count(text: str, default: int = 5) -> int:
    """Suy đoán số trang/slide từ câu lệnh (vd "...6 trang: ..."), có giới
    hạn hợp lý để tránh AI bị yêu cầu soạn quá dài."""
    match = re.search(r'(\d+)\s*(trang|slide)', text, flags=re.IGNORECASE)
    if not match:
        return default
    parsed_count = int(match.group(1))
    return max(1, min(parsed_count, 20))


@log_call
def suggest_file_name(user_text: str) -> str:
    """Gợi ý tên file từ câu lệnh — bỏ động từ/số trang, ghép các từ chính."""
    text = re.sub(r'\d+\s*(trang|slide)s?.*$', '', user_text, flags=re.IGNORECASE)
    words = re.findall(r"[^\W\d_]+", text)
    kept = [word for word in words if word.lower() not in _STOPWORDS] or words or ["TaiLieu"]
    name = "_".join(word[:1].upper() + word[1:] for word in kept[:6])
    name = re.sub(FORBIDDEN_FILENAME_CHARS, "", name).strip("_")
    return (name or "TaiLieu")[:60]
