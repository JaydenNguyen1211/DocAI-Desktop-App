"""Tham số tính lương + BHXH/BHYT/BHTN + thuế TNCN — TẤT CẢ đặt ở đây, KHÔNG
rải rác trong payroll.py, để người dùng tự rà soát/sửa khi quy định đổi.

⚠️ QUAN TRỌNG: các mức dưới đây là mức phổ biến áp dụng tại thời điểm viết
code (đầu năm 2026), lấy theo quy định đang có hiệu lực gần nhất mà tôi biết
được — GIỐNG NHƯ CHÍNH TÀI LIỆU KẾ HOẠCH ĐÃ LƯU Ý: "quy định thuế/bảo hiểm
thay đổi liên tục". Trước khi dùng cho việc tính lương thật, người dùng/kế
toán PHẢI tự đối chiếu lại các mức này với quy định hiện hành (Nghị định/
Nghị quyết/Thông tư mới nhất) — DocAI chỉ hỗ trợ tính toán theo tham số đã
khai báo, không tự cập nhật hay tự chịu trách nhiệm về tính đúng của mức.
"""
from dataclasses import dataclass, field

from ...logging_config import get_logger, log_call

logger = get_logger(__name__)


@dataclass
class InsuranceRates:
    """Tỷ lệ trích BHXH/BHYT/BHTN — áp trên "tiền lương tháng đóng BHXH".

    Nguồn tham chiếu tại thời điểm viết code: Luật BHXH 2024 (hiệu lực từ
    01/07/2025) + Luật Việc làm hiện hành cho BHTN — tỷ lệ % không đổi so với
    giai đoạn trước, chỉ đổi cách tính mức đóng tối đa với BHXH tự nguyện
    (không ảnh hưởng công thức MVP này).
    """
    # Người lao động (NLĐ) — trừ thẳng vào lương.
    bhxh_nld: float = 0.08
    bhyt_nld: float = 0.015
    bhtn_nld: float = 0.01
    # Doanh nghiệp (DN) — DN tự chi, không trừ lương NLĐ, chỉ để báo cáo chi phí.
    bhxh_dn: float = 0.175
    bhyt_dn: float = 0.03
    bhtn_dn: float = 0.01
    kpcd_dn: float = 0.02  # Kinh phí công đoàn — tính trên quỹ lương đóng BHXH.


# Mức lương cơ sở — trần đóng BHXH/BHYT = 20 × mức này.
# Nghị định 73/2024/NĐ-CP, hiệu lực 01/07/2024: 2.340.000đ.
LUONG_CO_SO = 2_340_000
BHXH_BHYT_TRAN = 20 * LUONG_CO_SO  # 46.800.000đ


# Mức lương tối thiểu vùng — trần đóng BHTN = 20 × mức vùng tương ứng.
# Nghị định 74/2024/NĐ-CP, hiệu lực 01/07/2024 (đơn vị: đồng/tháng).
LUONG_TOI_THIEU_VUNG = {
    "I": 4_960_000,
    "II": 4_410_000,
    "III": 3_860_000,
    "IV": 3_450_000,
}
_VUNG_ALIAS = {
    "1": "I", "I": "I", "VUNG I": "I", "VÙNG I": "I", "VUNG 1": "I", "VÙNG 1": "I",
    "2": "II", "II": "II", "VUNG II": "II", "VÙNG II": "II", "VUNG 2": "II", "VÙNG 2": "II",
    "3": "III", "III": "III", "VUNG III": "III", "VÙNG III": "III", "VUNG 3": "III", "VÙNG 3": "III",
    "4": "IV", "IV": "IV", "VUNG IV": "IV", "VÙNG IV": "IV", "VUNG 4": "IV", "VÙNG 4": "IV",
}


@log_call
def normalize_region(raw) -> str:
    """Chuẩn hoá ô "Vùng" trong bảng chấm công về "I"/"II"/"III"/"IV".

    Không nhận diện được → mặc định Vùng I (mức trần cao nhất, an toàn hơn
    là tính thiếu tiền đóng BHTN cho NLĐ)."""
    key = str(raw or "").strip().upper()
    return _VUNG_ALIAS.get(key, "I")


@log_call
def bhtn_tran(region: str) -> int:
    return 20 * LUONG_TOI_THIEU_VUNG[normalize_region(region)]


# Giảm trừ gia cảnh thuế TNCN — Nghị quyết 954/2020/UBTVQH14 (mức đang áp
# dụng tại thời điểm viết code; có thể đã có nghị quyết mới điều chỉnh mức
# này — CẦN TỰ KIỂM TRA LẠI).
GIAM_TRU_BAN_THAN = 11_000_000
GIAM_TRU_NGUOI_PHU_THUOC = 4_400_000


@dataclass
class TaxBracket:
    upto: float  # ngưỡng thu nhập tính thuế/tháng (đồng); inf = không giới hạn trên
    rate: float


# Biểu thuế TNCN lũy tiến từng phần — 7 bậc, theo Luật Thuế TNCN hiện hành.
TAX_BRACKETS: list[TaxBracket] = [
    TaxBracket(5_000_000, 0.05),
    TaxBracket(10_000_000, 0.10),
    TaxBracket(18_000_000, 0.15),
    TaxBracket(32_000_000, 0.20),
    TaxBracket(52_000_000, 0.25),
    TaxBracket(80_000_000, 0.30),
    TaxBracket(float("inf"), 0.35),
]


@dataclass
class PayrollRates:
    """Gộp TOÀN BỘ tham số ảnh hưởng kết quả tính lương vào 1 object — để có
    thể ghi đè nguyên khối bằng mức lấy từ server (`GET /rates`, xem
    `payroll._get_active_rates()`), thay vì chỉ ghi đè riêng % bảo hiểm.
    `version` là nhãn hiển thị cho người dùng biết đang dùng mức nào
    ("local-default" = mặc định cứng trong app, chưa từng đồng bộ được)."""
    insurance: InsuranceRates = field(default_factory=InsuranceRates)
    luong_co_so: int = LUONG_CO_SO
    luong_toi_thieu_vung: dict[str, int] = field(
        default_factory=lambda: dict(LUONG_TOI_THIEU_VUNG))
    giam_tru_ban_than: int = GIAM_TRU_BAN_THAN
    giam_tru_nguoi_phu_thuoc: int = GIAM_TRU_NGUOI_PHU_THUOC
    tax_brackets: list[TaxBracket] = field(default_factory=lambda: list(TAX_BRACKETS))
    version: str = "local-default"

    @log_call
    def bhxh_bhyt_tran(self) -> int:
        return 20 * self.luong_co_so

    @log_call
    def bhtn_tran(self, region: str) -> int:
        vung = self.luong_toi_thieu_vung.get(normalize_region(region))
        if vung is None:  # dữ liệu server thiếu vùng lạ → an toàn dùng Vùng I
            vung = self.luong_toi_thieu_vung.get("I", LUONG_TOI_THIEU_VUNG["I"])
        return 20 * vung


DEFAULT_RATES = InsuranceRates()
DEFAULT_PAYROLL_RATES = PayrollRates()
