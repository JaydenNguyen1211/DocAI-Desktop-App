"""Chuỗi văn bản tiếng Việt hiển thị cho người dùng — tập trung tại 1 nơi
thay vì rải rác trong từng file UI/module, để dễ rà soát cách hành văn, sửa
lỗi chính tả, hoặc sau này thêm đa ngôn ngữ mà không phải lục từng file.

Tổ chức: mỗi class là 1 "namespace" ứng với 1 màn hình/module (thường trùng
tên file hoặc tên class UI nơi chuỗi được dùng). Hằng số nào có phần thay đổi
được (tên file, số lượng…) viết dưới dạng template `str.format()` — nơi gọi
tự truyền giá trị, VD:

    from ...strings import Chat
    self.chat.add_system(Chat.OPENED_FILE.format(name=path_obj.name))

CHỈ chứa chuỗi hiển thị cho người dùng (label, tooltip, thông báo lỗi/trạng
thái). KHÔNG chứa: nội dung soạn sẵn chèn vào file Word/Excel/PPT do AI tạo
(đó là DỮ LIỆU xuất ra tài liệu, xem `modules.common.creators`/`business.payroll`),
prompt gửi AI (`modules.ai.generation`), hay các khóa dữ liệu/ánh xạ đã có sẵn
trong `app.constants` (EXT_MAP, FILE_BADGE…). Các thông báo lỗi validate đơn lẻ,
chỉ dùng ở đúng 1 nơi trong `modules/word|excel|pptx|image/ops.py` cũng CỐ Ý
để nguyên tại chỗ — gom vào đây không giảm trùng lặp, chỉ thêm tầng gián tiếp.
"""


# ══ UI dùng chung nhiều nơi ═══════════════════════════════════════════════

class Common:
    INPUT_PLACEHOLDER = "Nhập yêu cầu bằng tiếng Việt…"
    REMOVE_FROM_LIST = "Xóa khỏi danh sách"


# ══ app/ui/toast.py ═══════════════════════════════════════════════════════

class Toast:
    CACHE_HIT = "Cache HIT"
    CACHE_NEW = "Cache mới ghi"
    CACHE_NONE = "Không dùng cache"
    USAGE = "{headline} — đọc: {read:,} · ghi: {created:,} · mới: {fresh:,} token"


# ══ app/ui/empty_state.py ═════════════════════════════════════════════════

class EmptyState:
    TITLE = "Bạn muốn xử lý tài liệu gì?"
    SUBTITLE = "Gõ yêu cầu, hoặc kéo thả file vào cửa sổ."
    ATTACH_TOOLTIP = "Đính kèm file / thư mục"
    PICK_FILE_TITLE = "Chọn file"


# ══ app/ui/save_output.py ═════════════════════════════════════════════════

class SaveOutput:
    FILTER_DOCX = "Word (*.docx)"
    FILTER_XLSX = "Excel (*.xlsx)"
    FILTER_PDF = "PDF (*.pdf)"
    FILTER_PPTX = "PowerPoint (*.pptx)"
    FILTER_ALL = "Tất cả file (*.*)"
    DIALOG_TITLE = "Lưu file"


# ══ app/ui/doc_generation.py (DocGenWorker) ══════════════════════════════

class DocGeneration:
    UNKNOWN_ERROR = "Lỗi không xác định: {error}"
    NO_CONTENT = "AI không trả về nội dung để tạo file."
    BUILD_FAILED = "Không dựng được file: {error}"
    DEFAULT_SHEET_NAME = "Bảng tính"


# ══ app/ui/folder_tree.py ═════════════════════════════════════════════════

class FolderTree:
    FILTER_ALL = "Tất cả"


# ══ app/ui/chat.py ════════════════════════════════════════════════════════

class Chat:
    HEADER = "Chat AI"
    ATTACH_TOOLTIP = "Đính kèm file"
    STREAM_PLACEHOLDER = "…"


# ══ app/ui/doc_creation_widgets.py ════════════════════════════════════════

class DocCreation:
    CONFIRM_WITH_PAGES = "Tạo {label} ({page_count} trang)"
    CONFIRM = "Tạo {label}"
    UPGRADE_PRO = "Nâng cấp Pro"
    LATER = "Để sau"
    GEN_FAILED_TITLE = "Tạo file thất bại"
    GEN_PROGRESS = "trang {done}/{total}"
    RETRY = "Thử lại"
    SAVE_PARTIAL = "Lưu {done} trang đã có"


# ══ app/ui/preview.py ═════════════════════════════════════════════════════

class Preview:
    OPEN_FOLDER = "Mở thư mục"
    OPEN_WITH = "Mở bằng…"
    OPEN_WITH_TYPE = "Mở bằng {label}"
    OPEN_WITH_DEFAULT_LABEL = "ứng dụng"
    EXTRACT_TEXT = "Trích xuất văn bản"
    CONVERT = "Chuyển đổi định dạng"
    SEARCH_PLACEHOLDER = "Tìm từ khóa trong PDF…"
    FOOTER = "Xem trước tài liệu — giữ nguyên format"
    NO_PAGES = "Không có trang nào."
    UNSUPPORTED_FORMAT = "Định dạng chưa hỗ trợ xem trước."
    NEEDS_DEPENDENCIES = "Cần cài PyMuPDF + Pillow để xem trước tài liệu."
    OPENING = "Đang mở tài liệu…"
    PREVIEW_FAILED = "Không xem trước được:\n{error}"
    NOT_FOUND = "Không tìm thấy"
    SEARCH_RESULT = "Trang {page} ({pos}/{total})"


# ══ app/ui/sidebar.py ═════════════════════════════════════════════════════

class Sidebar:
    TAB_FILE = "Tệp"
    TAB_FOLDER = "Thư mục"
    NEW_CHAT = "+  Trò chuyện mới"
    RECENT_LABEL_BASE = "Gần đây"   # hiển thị viết hoa toàn bộ ở nơi gọi
    OPEN_FOLDER_BTN = "  Mở thư mục làm việc"
    FOLDER_EMPTY_HINT = (
        "Chưa mở thư mục nào.\n"
        "Mở một thư mục rồi nói chuyện với AI về các file trong đó.")
    SCANNING = "Đang quét thư mục…"
    SCANNING_PROGRESS = "Đang quét thư mục…\n{done}{suffix} file"
    CHANGE = "Đổi"
    RECENT_EMPTY = "Chưa có cuộc nào"
    PLAN_LABEL = "Gói {plan}"
    PLAN_HINT = "tác vụ AI còn lại tháng này"
    FOLDER_FILE_COUNT = "{count} file"


# ══ app/ui/splash.py ══════════════════════════════════════════════════════

class Splash:
    WINDOW_TITLE = "{app_name} — Đăng nhập"

    WORD_FOUND = "Đã tìm thấy Microsoft Word — dùng Word COM để xử lý tài liệu."
    WORD_NOT_FOUND_LIBREOFFICE = "Không thấy Microsoft Word — sẽ dùng LibreOffice dự phòng."
    WORD_NOT_FOUND_NONE = (
        "Không thấy Microsoft Word hay LibreOffice — chức năng chuyển đổi bị hạn chế.")
    OFFICE_MAC_FOUND = (
        "Đã tìm thấy Microsoft Office for Mac — dùng Office để xử lý tài liệu.")
    OFFICE_MAC_NOT_FOUND_LIBREOFFICE = "Không thấy Microsoft Office — sẽ dùng LibreOffice dự phòng."
    OFFICE_MAC_NOT_FOUND_NONE = (
        "Không thấy Microsoft Office hay LibreOffice — chức năng chuyển đổi bị hạn chế.")
    LIBREOFFICE_LINUX = "Đang dùng LibreOffice để xử lý tài liệu."
    LIBREOFFICE_LINUX_NOT_FOUND = "Không thấy LibreOffice — chức năng chuyển đổi bị hạn chế."

    EMAIL_LABEL = "Email"
    EMAIL_PLACEHOLDER = "ban@congty.vn"
    PASSWORD_LABEL = "Mật khẩu"
    PASSWORD_PLACEHOLDER = "Tối thiểu 6 ký tự"
    LOGIN = "Đăng nhập"
    SIGNUP = "Đăng ký"
    NO_ACCOUNT_HINT = "Chưa có tài khoản?"
    HAS_ACCOUNT_HINT = "Đã có tài khoản?"
    CHECKING_OFFICE = "Đang kiểm tra phần mềm Office…"
    LOGGING_IN = "Đang đăng nhập…"
    SIGNING_UP = "Đang đăng ký…"
    SERVER_NOT_CONFIGURED = (
        "Chưa cấu hình server: điền FIREBASE_API_KEY và API_BASE_URL "
        "trong server_config.py (hoặc config.json).")
    INVALID_EMAIL = "Email không hợp lệ."
    PASSWORD_TOO_SHORT = "Mật khẩu cần tối thiểu 6 ký tự."


# ══ app/ui/modals.py ══════════════════════════════════════════════════════

class SettingsDialog:
    TITLE = "Cài đặt & tài khoản"
    TAB_ACCOUNT = "Tài khoản"
    TAB_BUSINESS = "Doanh nghiệp"
    TAB_LIMITS = "Giới hạn"
    SAVE = "Lưu thay đổi"
    FOOTNOTE = "* Thông tin doanh nghiệp dùng để AI tự điền khi soạn văn bản hành chính"
    FIELD_ACCOUNT = "Tài khoản"
    FIELD_PLAN = "Gói hiện tại"
    FIELD_QUOTA = "Quota còn lại"
    LOADING = "Đang tải…"
    LOGOUT = "Đăng xuất"
    PLAN_PRO = "Pro"
    PLAN_FREE = "Free"
    QUOTA_VALUE = "{remaining} / {limit} lượt"
    FIELD_COMPANY_NAME = "Tên doanh nghiệp"
    FIELD_TAX_CODE = "Mã số thuế (MST)"
    FIELD_ADDRESS = "Địa chỉ"
    FIELD_REPRESENTATIVE = "Người đại diện"
    LIMITS_PLACEHOLDER = "Chi tiết quota theo gói sẽ có ở Phase 2."


class ConvertDialog:
    TITLE = "Chuyển đổi định dạng"
    TARGET_FORMAT_LABEL = "Định dạng đích"
    NO_EXTRA_IMAGES = "Chưa chọn thêm ảnh nào"
    ADD_EXTRA_IMAGE = "+ Thêm ảnh khác…"
    EXTRA_IMAGES_ADDED = "Đã chọn thêm {count} ảnh (gộp theo thứ tự)"
    CANCEL = "Hủy"
    CONVERT = "Chuyển đổi"
    CONVERTING = "Đang chuyển đổi…"
    PICK_EXTRA_TITLE = "Chọn thêm ảnh để gộp"
    IMAGE_FILTER = "Ảnh (*.png *.jpg *.jpeg *.webp *.bmp)"
    ERROR_PREFIX = "⚠ {message}"


class OverwriteDialog:
    TITLE = "Xác nhận ghi đè"
    HEADING = "File sẽ bị ghi đè"
    MESSAGE = "Bạn có muốn lưu một bản sao trước khi\nAI ghi đè lên «{file_name}» không?"
    SAVE_COPY = "Lưu bản sao"
    OVERWRITE = "Ghi đè"


class NewDocumentDialog:
    DEFAULT_TITLE = "Tạo tài liệu mới"
    FILE_NAME_LABEL = "Tên file"
    NAME_EMPTY = "Tên file không được để trống"
    NAME_FORBIDDEN = "Tên file không được chứa {chars}"
    SAVE_LOCATION_LABEL = "Lưu vào"
    CHANGE = "Đổi…"
    OPEN_AFTER_CREATE = "Mở file để làm việc ngay sau khi tạo"
    DUPLICATE_WARNING = "⚠ Đã tồn tại file cùng tên trong thư mục này"
    OVERWRITE_OLD = "Ghi đè file cũ"
    RENAME_TO = "Đổi tên thành {name}"
    CANCEL = "Hủy"
    CREATE_FILE = "Tạo file"
    PATH_ERROR_TITLE = "Không thể lưu vào thư mục này"
    PATH_ERROR_MSG = "{dir} không còn kết nối, hoặc ứng dụng không có quyền ghi vào đây."
    CHOOSE_OTHER_LOCATION = "Chọn nơi lưu khác"
    RETRY = "Thử lại"
    PICK_SAVE_DIR_TITLE = "Chọn nơi lưu"


# ══ app/ui/main_window.py ═════════════════════════════════════════════════

class MainWindow:
    # Mở file / kéo-thả
    PLAN_LABEL = "Gói {plan}"
    PLAN_LABEL_WITH_QUOTA = "Gói {plan} · {remaining}/{limit} lượt"
    FILE_NOT_FOUND = "File không tồn tại:\n{path}"
    FORMAT_UNSUPPORTED = "Định dạng chưa hỗ trợ: {ext}"
    ATTACHED_TO_CONTEXT = "Đã đính kèm «{name}» vào ngữ cảnh"
    OPENED_FILE = "Đã mở «{name}»"
    ATTACHED_FILE = "Đã đính kèm «{name}»"
    PICK_FILE_TITLE = "Chọn file"

    # Chế độ Thư mục
    PICK_FOLDER_TITLE = "Chọn thư mục làm việc"
    FOLDER_OPENED = "Đã mở thư mục «{name}» — {count} file"
    CLARIFY_CREATE = (
        "Bạn muốn tạo file loại gì (Word/Excel/PowerPoint) và tên là gì? "
        'VD: Tạo file Excel mới tên "BaoCao.xlsx"')
    CLARIFY_DELETE = (
        "Chưa rõ bạn muốn xóa file nào — hãy nêu rõ tên file (có thể để trong ngoặc kép).")
    CLARIFY_RENAME = (
        'Chưa rõ file cần đổi tên hoặc tên mới — VD: Đổi tên "A.docx" thành "B.docx"')
    CLARIFY_OPEN = "Chưa rõ bạn muốn mở file nào trong thư mục — hãy nêu rõ tên file."
    CLARIFY_GENERIC = "Bạn nói rõ hơn giúp mình nhé."
    CREATED_FILE = "✓ Đã tạo «{name}» trong thư mục"
    JUST_CREATED_NOTE = "Vừa tạo"
    DELETED_FILE = "✓ Đã xóa «{name}» khỏi thư mục"
    UNDO = "Hoàn tác"
    RESTORED_FILE = "✓ Đã khôi phục «{name}»"
    RENAMED_FILE = "✓ Đã đổi tên thành «{name}» — nội dung giữ nguyên"

    ERROR_PREFIX = "⚠ {message}"
    SUGGESTIONS_IN_CONTEXT = "File trong ngữ cảnh:"

    # Chat AI / streaming
    PROCESSING = "Đang xử lý…"
    NO_ANALYSIS_CONTENT = "(Không có nội dung phân tích)"
    NO_REPLY_CONTENT = "(Không có nội dung trả về)"
    FILE_UPDATED = "Đã cập nhật file."

    # Tính lương
    CALCULATING_PAYROLL = "Đang tính lương…"
    PAYROLL_SUMMARY = (
        "Đã tính lương cho {count} nhân viên.\n\n"
        "· Tổng lương thực nhận (Net): {net:,.0f} đ\n"
        "· Tổng BHXH/BHYT/BHTN người lao động trích: {nld:,.0f} đ\n"
        "· Tổng chi phí doanh nghiệp (lương + bảo hiểm DN đóng): {dn:,.0f} đ\n\n"
        "Bấm vào file bên dưới để xem trước & lưu — sheet «Chi phí DN» có chi tiết "
        "phần doanh nghiệp phải đóng.\n\n"
        "⚠ Số liệu tính theo mức BHXH/BHYT/BHTN/thuế TNCN «{rates_version}» — "
        "nhờ kế toán đối chiếu lại trước khi dùng chính thức."
    )

    # Sau khi AI trả lời / sửa file
    EDIT_NOTE = "✓ «{name}» — {note}"
    STAGED_FILE_NOTE = "Bấm để xem trước & lưu"
    OVERWRITTEN = "✓ Đã ghi đè thay đổi lên «{name}»"
    COPY_KEPT_ORIGINAL = "✓ Đã lưu bản sao, giữ nguyên file gốc"

    # Chuyển đổi định dạng
    EXPORTED_IMAGES = "✓ Đã xuất {count} ảnh vào thư mục «{name}»"
    CONVERT_DONE = "✓ Chuyển đổi định dạng hoàn tất"

    # Trích xuất văn bản PDF
    EXTRACT_FAILED = "⚠ Không trích xuất được văn bản: {error}"
    SAVE_EXTRACTED_TITLE = "Lưu văn bản đã trích xuất"
    TEXT_FILTER = "Văn bản (*.txt)"
    WRITE_FAILED = "⚠ Không ghi được file: {error}"
    EXTRACTED_TO_FILE = "✓ Đã trích xuất văn bản vào «{name}»"

    # Tạo tài liệu mới bằng chat
    AMBIGUOUS_CREATE_QUESTION = (
        "Bạn cho mình biết thêm để soạn đúng ý nhé:\n"
        "· Nội dung cụ thể là gì? (vd: doanh thu, hợp đồng, giới thiệu…)\n"
        "· Dạng văn bản (Word), bảng biểu (Excel) hay trình bày (PowerPoint)?"
    )
    AMBIGUOUS_CREATE_CHIP_1 = "Báo cáo doanh thu · Excel"
    AMBIGUOUS_CREATE_CHIP_2 = "Báo cáo tiến độ · Word"
    QUOTA_EXCEEDED = "Bạn đã dùng hết {limit} tác vụ AI miễn phí tháng này"
    QUOTA_HINT = "Nâng cấp Pro để tiếp tục tạo file ngay, hoặc chờ tới kỳ làm mới."
    ASK_FORMAT = "Mình sẽ tạo giúp bạn. Xuất ra định dạng nào?"
    GENERATING_PLAIN = "Đang soạn…"
    GENERATING = "Đang soạn «{name}»…"
    GENERATING_PAGE = "Đang soạn {label} ({page}/{total})…"
    GEN_DONE = "Đã tạo xong. Bạn xem trước rồi lưu, hoặc bảo mình chỉnh tiếp."
    GEN_FAILED_PREVIEW = "Tạo file thất bại:\n{message}"
    CANNOT_SAVE_PARTIAL_EXCEL = "⚠ Chưa thể lưu một phần cho file Excel."
    SAVE_PARTIAL_FAILED = "⚠ Không lưu được phần đã có: {error}"
    PARTIAL_SAVED = "Đã lưu tạm {count} phần đã soạn được. Bạn xem trước rồi lưu."
    SAVE_FILE_TYPE = "Lưu file {label}"
    SAVE_LABELS = {"word": "Word", "excel": "Excel", "ppt": "PowerPoint"}
    SAVE_FAILED = "⚠ Không lưu được file: {error}"
    SAVED_AND_READY = "✓ Đã lưu «{name}» vào {dir}. Bạn muốn chỉnh gì tiếp không?"

    # File card (kết quả mock chưa tích hợp)
    MOCK_FILE_RESULT = "«{name}»\n\nFile kết quả sẽ được tạo thật khi tích hợp backend AI."
    SAVED_CONTINUE = "✓ Đã lưu «{name}» — bạn có thể tiếp tục ra lệnh trên file này"


# ══ account/api_client.py ═════════════════════════════════════════════════

class Account:
    # Ánh xạ mã lỗi Firebase → thông báo tiếng Việt.
    ERROR_MAP = {
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
    WEAK_PASSWORD = "Mật khẩu quá yếu (cần tối thiểu 6 ký tự)."
    AUTH_ERROR = "Lỗi xác thực: {code}"
    UNKNOWN_ERROR = "Lỗi không rõ."
    INVALID_SESSION = "Phiên đăng nhập không hợp lệ, vui lòng đăng nhập lại."
    NOT_LOGGED_IN = "Chưa đăng nhập."
    NO_CONNECTION = "Không kết nối được máy chủ. Kiểm tra mạng."
    SESSION_EXPIRED = "Phiên đăng nhập hết hạn, vui lòng đăng nhập lại."
    QUOTA_EXCEEDED = "Đã hết lượt gọi AI trong tháng. Nâng cấp gói để tiếp tục."
    SERVER_ERROR = "Máy chủ báo lỗi ({status})."


# ══ modules/common/converters.py ══════════════════════════════════════════

class Converters:
    WORD_TO_PDF_FAILED = "Không chuyển được Word sang PDF: {error}"
    PPT_TO_PDF_FAILED = "Không chuyển được PowerPoint sang PDF: {error}"
    PDF_NO_TEXT_LAYER = (
        "PDF này không có lớp chữ (có thể là bản quét ảnh) — không tự "
        "chuyển sang Word được. Hãy dùng OCR ảnh (mở từng trang dạng ảnh) "
        "thay vì chức năng này.")
    WRITE_FAILED_FILE_OPEN = "Không ghi được «{name}» — file đích đang mở ở chương trình khác."
    NO_TABLE_DATA_FROM_FILE = "AI không trích xuất được dữ liệu dạng bảng từ file này."
    PDF_NO_PAGES = "PDF không có trang nào để xuất."
    NO_IMAGES_TO_CONVERT = "Không có ảnh nào để chuyển đổi."
    IMAGE_TO_PDF_FAILED = "Không tạo được PDF từ ảnh: {error}"
    NO_TEXT_READ_FROM_IMAGES = "AI không đọc được nội dung văn bản từ (các) ảnh này."
    SEARCHABLE_PDF_FAILED = "Không tạo được PDF có thể tìm kiếm: {error}"
    NO_TABLE_DATA_FROM_IMAGE = "AI không trích xuất được dữ liệu dạng bảng từ ảnh này."
    TARGET_FORMAT_UNSUPPORTED = "Định dạng đích chưa hỗ trợ: {ext}."
    IMAGE_FORMAT_CONVERT_FAILED = "Không chuyển được định dạng ảnh: {error}"
    CONVERSION_UNSUPPORTED = "Chưa hỗ trợ chuyển từ {source_type} sang {target_ext}."


# ══ modules/common/folder_ops.py ══════════════════════════════════════════

class FolderOps:
    FILE_ALREADY_EXISTS = 'File "{name}" đã tồn tại trong thư mục.'
    FILE_TYPE_UNSUPPORTED = "Loại file chưa hỗ trợ: {ftype}"
    CREATE_FAILED = "Không tạo được file: {error}"
    DELETE_FAILED = "Không xóa được file: {error}"
    RESTORE_NAME_CONFLICT = "Không khôi phục được — đã có file khác cùng tên tại vị trí cũ."
    RESTORE_FAILED = "Không khôi phục được: {error}"
    RENAME_NAME_CONFLICT = 'Đã có file tên "{name}" trong thư mục.'
    RENAME_FAILED = "Không đổi được tên file: {error}"


# ══ modules/common/doc_set.py ═════════════════════════════════════════════

class DocSet:
    FORMAT_UNSUPPORTED = "Định dạng chưa hỗ trợ: {ext_or_path}"
    NO_READABLE_CONTENT = "Không đọc được nội dung file nào trong danh sách đã chọn."
    MERGE_FAILED = "Không gộp được file: {error}"
    NO_FILES_SELECTED = "Chưa chọn file nào để xử lý."
    FILES_NOT_FOUND = "Không tìm thấy file: {names}"
    CLARIFY_WHICH_FILES = (
        "Chưa rõ bạn muốn xử lý (các) file nào — hãy chọn file trong danh sách "
        "hoặc nói rõ tên file.")
    CLARIFY_MORE_FILES = (
        "Bạn nhắc tới {hint_count} file nhưng mình mới nhận diện được "
        "{resolved_count}. Còn thiếu file nào nữa?")


# ══ modules/common/creators.py ════════════════════════════════════════════

class Creators:
    MISSING_DOCX_LIB = "Thiếu thư viện python-docx để tạo file Word."
    MISSING_PPTX_LIB = "Thiếu thư viện python-pptx để tạo file PowerPoint."


# ══ modules/common/_soffice.py / _msoffice_mac.py ═════════════════════════

class Office:
    LIBREOFFICE_NOT_FOUND = (
        "Không tìm thấy LibreOffice. Hãy cài LibreOffice (https://libreoffice.org) "
        "hoặc Microsoft Office để chuyển đổi / xem trước file Word, PowerPoint, Excel.")
    LIBREOFFICE_ERROR = "LibreOffice báo lỗi: {detail}"
    LIBREOFFICE_NO_PDF = "LibreOffice không tạo được file PDF."
    APPLESCRIPT_FAILED = "AppleScript thất bại không rõ nguyên nhân."
    INVALID_PATH_QUOTE = "Đường dẫn chứa ký tự không hợp lệ (dấu nháy kép): {path}"
    MISSING_DOCX2PDF = "Cần cài docx2pdf: pip install docx2pdf"
    DOCX2PDF_FAILED = "docx2pdf thất bại: {error}"


# ══ modules/excel/com.py ══════════════════════════════════════════════════

class ExcelCom:
    MISSING_EXCEL = "Cần cài Microsoft Excel trên máy để chuyển đổi sang PDF."
    EXPORT_FAILED = "Không xuất được PDF từ Excel: {error}"


# ══ modules/business/payroll.py ═══════════════════════════════════════════

class Payroll:
    MISSING_COLUMNS = (
        "Không tìm thấy cột bắt buộc: {names}.\n"
        "Bảng chấm công cần có dòng tiêu đề (dòng 1) với các cột (không cần đủ hết, "
        "chỉ 2 cột đầu là bắt buộc):\n{expected}")
    READ_FAILED = "Không đọc được file «{name}»: {error}"
    FILE_EMPTY = "File «{name}» trống, không có dữ liệu."
    NO_VALID_EMPLOYEES = "Không có dòng nhân viên nào hợp lệ trong «{name}» (mỗi dòng cần có Họ tên)."
    SALARY_MUST_BE_POSITIVE = "«{name}»: lương cơ bản phải lớn hơn 0."
    STANDARD_DAYS_MUST_BE_POSITIVE = "«{name}»: ngày công chuẩn phải lớn hơn 0."
    OUTPUT_FILE_LOCKED = (
        "Không ghi được file kết quả — «{name}» đang mở ở "
        "chương trình khác, hãy đóng lại rồi thử lại.")


# ══ pipeline/editing.py ═══════════════════════════════════════════════════

class Editing:
    TYPE_NOT_EDITABLE = "Bản này chỉ sửa được file Word, Excel, PowerPoint và ảnh."
    FORMAT_NOT_EDITABLE = (
        "Chưa sửa được định dạng {suffix} — hãy lưu lại thành {target_ext} rồi mở lại.")
    MISSING_DOCX_LIB = "Thiếu thư viện python-docx để sửa file Word."
    MISSING_OPENPYXL_LIB = "Thiếu thư viện openpyxl để sửa file Excel."
    MISSING_PPTX_LIB = "Thiếu thư viện python-pptx để sửa file PowerPoint."
    OUTLINE_READ_FAILED = "Không đọc được cấu trúc file: {error}"
