"""AI — nội dung câu trả lời lấy từ server (Claude), còn module này lo phần
phân loại ý định để giữ UX file-card / ghi đè cho bản MVP.

`build_reply()` (mock cũ) vẫn giữ để chạy offline khi cần; luồng thật dùng
`classify_intent()` kết hợp text trả về từ server.
"""
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from ...logging_config import get_logger, log_call

logger = get_logger(__name__)


@dataclass
class MockReply:
    text: str
    artifact: Optional[str] = None   # tên file kết quả đính kèm (card trong chat)
    is_edit: bool = False            # True → app hỏi ghi đè / lưu bản sao
    edit_keywords: list = field(default_factory=list)


_EDIT_WORDS = ("sửa", "chỉnh", "thay", "đổi", "thêm", "xóa", "xoá", "cập nhật", "viết lại")


@log_call
def classify_intent(user_text: str, file_name: Optional[str],
                    file_type: Optional[str]) -> tuple[Optional[str], bool]:
    """Suy ra (tên file kết quả, có phải yêu cầu chỉnh sửa) từ câu lệnh.

    Dùng để hiển thị card file đính kèm và modal xác nhận ghi đè — nội dung
    văn bản thực tế do server (Claude) trả về.
    """
    low = user_text.lower()
    artifact: Optional[str] = None

    if "trích xuất" in low or "hóa đơn" in low or "hoá đơn" in low:
        artifact = "Bang_ke_hoa_don.xlsx"
    elif "lương" in low or "chấm công" in low:
        artifact = "Bang_luong.xlsx"
    elif "thuế" in low or "kê khai" in low or "gtgt" in low or "tncn" in low:
        artifact = "Bang_ke_khai_thue.xlsx"
    elif ("công văn" in low or "tờ trình" in low or "quyết định" in low
          or "thông báo" in low or "soạn" in low):
        artifact = "Cong_van_moi.docx"

    is_edit = bool(file_name) and any(word in low for word in _EDIT_WORDS)
    return artifact, is_edit


@log_call
def build_reply(user_text: str, file_name: Optional[str],
                file_type: Optional[str], business: dict) -> MockReply:
    low = user_text.lower()
    company = business.get("company_name") or "Công ty của bạn"

    if "tóm tắt" in low:
        target = f" «{file_name}»" if file_name else ""
        return MockReply(
            f"Đã tóm tắt tài liệu{target}:\n\n"
            "• Nội dung chính: hợp đồng cung cấp dịch vụ giữa hai bên, "
            "thời hạn 12 tháng kể từ ngày ký.\n"
            "• Giá trị hợp đồng: 250.000.000 ₫ (chưa gồm VAT 10%).\n"
            "• Điều khoản thanh toán: 2 đợt — 50% tạm ứng, 50% khi nghiệm thu.\n"
            "• Lưu ý: điều 7.2 quy định phạt chậm tiến độ 0,05%/ngày.\n\n"
            "Bạn cần tôi soạn văn bản phản hồi hay trích xuất điều khoản nào chi tiết hơn không?"
        )

    if "trích xuất" in low or "hóa đơn" in low or "hoá đơn" in low:
        return MockReply(
            "Đã đọc và trích xuất 8 dòng hóa đơn từ tài liệu.\n\n"
            "Phân loại: 6 hóa đơn đầu vào (mua hàng hóa, dịch vụ) · 2 hóa đơn đầu ra.\n"
            "Tổng giá trị trước thuế: 182.400.000 ₫ · VAT: 18.240.000 ₫.\n\n"
            "Đang tạo bảng kê Excel theo mẫu kê khai GTGT…",
            artifact="Bang_ke_hoa_don.xlsx",
        )

    if "lương" in low or "chấm công" in low:
        return MockReply(
            "Đã tính bảng lương từ dữ liệu chấm công:\n\n"
            "• 12 nhân sự · 26 công chuẩn · 3 người có tăng ca.\n"
            "• Tổng lương gross: 186.500.000 ₫ → net: 164.720.000 ₫.\n"
            "• Trích BHXH 8% + BHYT 1,5% + BHTN 1% (NLĐ) — "
            "phần doanh nghiệp 21,5% đã tách cột riêng.\n"
            "• Kinh phí công đoàn 2% tính trên quỹ lương đóng BHXH.\n\n"
            "File bảng lương chi tiết đính kèm bên dưới.",
            artifact="Bang_luong_T7.xlsx",
        )

    if "công văn" in low or "tờ trình" in low or "quyết định" in low or "thông báo" in low or "soạn" in low:
        return MockReply(
            f"Đã soạn công văn theo đúng thể thức hành chính "
            f"(Nghị định 30/2020/NĐ-CP):\n\n"
            f"• Đơn vị ban hành: {company}\n"
            "• Quốc hiệu, tiêu ngữ, số/ký hiệu, ngày tháng đầy đủ.\n"
            "• Nơi nhận và chữ ký người đại diện lấy từ hồ sơ doanh nghiệp "
            "trong Cài đặt.\n\n"
            "Bạn xem bản nháp đính kèm và cho tôi biết cần điều chỉnh gì nhé.",
            artifact="Cong_van_moi.docx",
        )

    if "thuế" in low or "kê khai" in low or "gtgt" in low or "tncn" in low:
        return MockReply(
            "Đã dựng bảng kê khai thuế từ dữ liệu hóa đơn:\n\n"
            "• Bảng kê GTGT đầu vào – đầu ra theo mẫu 01/GTGT.\n"
            "• So sánh 2 phương án: khấu trừ vs trực tiếp — "
            "phương án khấu trừ lợi hơn 4.360.000 ₫ trong kỳ này.\n\n"
            "File Excel chi tiết đính kèm bên dưới.",
            artifact="Bang_ke_khai_thue.xlsx",
        )

    if file_name and any(word in low for word in _EDIT_WORDS):
        return MockReply(
            f"Đã thực hiện chỉnh sửa trên «{file_name}» theo yêu cầu:\n\n"
            "• Giữ nguyên toàn bộ style, font, bảng biểu và header/footer.\n"
            "• Các thay đổi đã áp vào đúng vị trí trong tài liệu.\n\n"
            "Tài liệu đã sẵn sàng để lưu.",
            is_edit=True,
        )

    if file_name:
        return MockReply(
            f"Tôi đang mở «{file_name}». Bạn có thể yêu cầu tôi:\n\n"
            "• Tóm tắt nội dung chính\n"
            "• Chỉnh sửa trực tiếp (giữ nguyên format)\n"
            "• Trích xuất dữ liệu thành bảng Excel\n"
            "• Chuyển đổi định dạng\n\n"
            "Cứ gõ yêu cầu bằng tiếng Việt tự nhiên nhé!"
        )

    return MockReply(
        "Tôi là DocAI — trợ lý xử lý tài liệu văn phòng. Tôi có thể:\n\n"
        "• Soạn công văn, tờ trình, quyết định đúng thể thức\n"
        "• Tính bảng lương kèm BHXH/BHYT/BHTN từ chấm công\n"
        "• Đọc hóa đơn điện tử và dựng bảng kê thuế\n"
        "• Tóm tắt, chỉnh sửa, chuyển đổi tài liệu\n\n"
        "Kéo thả file vào panel bên trái hoặc gõ yêu cầu ngay tại đây."
    )


@log_call
def suggest_save_name(file_name: str) -> str:
    path = Path(file_name)
    return f"{path.stem}_ban_sao{path.suffix}"
