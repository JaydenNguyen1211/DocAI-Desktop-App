"""Parameters for computing payroll + social/health/unemployment insurance +
personal income tax — ALL kept here, NOT scattered across payroll.py, so
users can review/edit them themselves when regulations change.

⚠️ IMPORTANT: the rates below are the common values in effect at the time
this code was written (early 2026), based on the most recent regulations I'm
aware of — JUST AS THE PLANNING DOCS THEMSELVES NOTE: "tax/insurance
regulations change constantly". Before using this for real payroll runs,
the user/accountant MUST cross-check these rates against current regulations
(the latest Decree/Resolution/Circular) — DocAI only assists with
computation based on the declared parameters, it does not auto-update or
take responsibility for the rates' accuracy.
"""
from dataclasses import dataclass, field

from ...logging_config import get_logger, log_call

logger = get_logger(__name__)


@dataclass
class InsuranceRates:
    """Social/health/unemployment insurance contribution rates — applied to
    the "monthly salary used for social insurance contribution".

    Reference source at the time this code was written: Social Insurance Law
    2024 (effective 07/01/2025) + the current Employment Law for
    unemployment insurance — the % rates are unchanged from the previous
    period, only the maximum contribution calculation for voluntary social
    insurance changed (doesn't affect this MVP's formula).
    """
    # Employee side (NLĐ) — withheld directly from salary.
    bhxh_nld: float = 0.08
    bhyt_nld: float = 0.015
    bhtn_nld: float = 0.01
    # Employer side (DN) — paid by the employer, not withheld from the
    # employee's salary, only reported as cost.
    bhxh_dn: float = 0.175
    bhyt_dn: float = 0.03
    bhtn_dn: float = 0.01
    kpcd_dn: float = 0.02  # Trade union fee — computed on the social-insurance salary fund.


# Base salary — the social/health insurance contribution ceiling = 20 × this rate.
# Decree 73/2024/NĐ-CP, effective 07/01/2024: 2,340,000 VND.
LUONG_CO_SO = 2_340_000
BHXH_BHYT_TRAN = 20 * LUONG_CO_SO  # 46,800,000 VND


# Regional minimum wage — the unemployment insurance contribution ceiling =
# 20 × the corresponding regional rate.
# Decree 74/2024/NĐ-CP, effective 07/01/2024 (unit: VND/month).
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
    """Normalize the "Region" cell in the timesheet to "I"/"II"/"III"/"IV".

    Unrecognized → defaults to Region I (the highest ceiling, safer than
    under-computing the employee's unemployment insurance contribution)."""
    key = str(raw or "").strip().upper()
    return _VUNG_ALIAS.get(key, "I")


@log_call
def bhtn_tran(region: str) -> int:
    return 20 * LUONG_TOI_THIEU_VUNG[normalize_region(region)]


# Personal income tax family deduction — Resolution 954/2020/UBTVQH14 (the
# rate in effect at the time this code was written; a newer resolution may
# have adjusted this — NEEDS TO BE VERIFIED).
GIAM_TRU_BAN_THAN = 11_000_000
GIAM_TRU_NGUOI_PHU_THUOC = 4_400_000


@dataclass
class TaxBracket:
    upto: float  # taxable income threshold/month (VND); inf = no upper limit
    rate: float


# Progressive personal income tax brackets — 7 tiers, per the current PIT Law.
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
    """Bundles EVERY parameter that affects the payroll result into 1 object
    — so it can be overridden wholesale by rates fetched from the server
    (`GET /rates`, see `payroll._get_active_rates()`), instead of only
    overriding the insurance %. `version` is a label shown to the user
    indicating which rates are in use ("local-default" = the app's hardcoded
    default, never successfully synced)."""
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
        if vung is None:  # server data missing an unfamiliar region → safely default to Region I
            vung = self.luong_toi_thieu_vung.get("I", LUONG_TOI_THIEU_VUNG["I"])
        return 20 * vung


DEFAULT_RATES = InsuranceRates()
DEFAULT_PAYROLL_RATES = PayrollRates()
