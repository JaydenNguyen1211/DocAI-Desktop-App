APP_NAME = "DocAI"
APP_TAGLINE = "Trợ lý AI xử lý tài liệu văn phòng"

# Gói & quota (mock — backend Phase 2)
PLAN_NAME = "Pro"
PLAN_QUOTA = 500

# Nhận diện loại file
EXT_MAP = {
    ".docx": "word", ".doc": "word",
    ".xlsx": "excel", ".xls": "excel",
    ".pptx": "ppt", ".ppt": "ppt",
    ".pdf": "pdf",
    ".png": "image", ".jpg": "image", ".jpeg": "image",
    ".bmp": "image", ".gif": "image", ".webp": "image",
}

FILE_BADGE = {
    "word": "DOC", "excel": "XLS", "ppt": "PPT", "pdf": "PDF", "image": "IMG",
}

FILE_DIALOG_FILTER = (
    "Tài liệu (*.docx *.doc *.xlsx *.xls *.pptx *.ppt *.pdf *.png *.jpg *.jpeg *.bmp *.webp);;"
    "Word (*.docx *.doc);;Excel (*.xlsx *.xls);;PowerPoint (*.pptx *.ppt);;"
    "PDF (*.pdf);;Ảnh (*.png *.jpg *.jpeg *.bmp *.webp)"
)

# Chip gợi ý ở chat AI trung tâm (nhãn · màu icon ◆) — theo thiết kế
EMPTY_CHIPS = [
    ("Tóm tắt hợp đồng dài",       "#C0392B"),
    ("Trích xuất hóa đơn → Excel", "#1F8A5B"),
    ("Tính lương + BHXH",          "#2A6FDB"),
    ("Soạn công văn hành chính",   "#6E6A63"),
]

# Badge loại file cho item "Gần đây" trong sidebar: (nhãn, khóa màu)
FILE_BADGE_STYLE = {
    "word":  ("DOC", "doc"),
    "excel": ("XLS", "xls"),
    "pdf":   ("PDF", "pdf"),
    "ppt":   ("PPT", "other"),
    "image": ("IMG", "other"),
    "xml":   ("XML", "other"),
}

# Chế độ Thư mục: nhận diện thêm XML (hóa đơn điện tử) — chỉ dùng để tick
# chọn làm ngữ cảnh, không cần renderer xem trước như chế độ Tệp.
FOLDER_EXT_MAP = {**EXT_MAP, ".xml": "xml"}
FOLDER_BADGE = {**FILE_BADGE, "xml": "XML"}

# Quick-action chips theo ngữ cảnh file đang mở
CONTEXT_CHIPS = {
    "word":  ["Tóm tắt", "Soạn công văn", "Chuyển sang PDF"],
    "excel": ["Tính lương", "Tạo bảng kê thuế", "Chuyển sang PDF"],
    "ppt":   ["Tóm tắt", "Chuyển sang PDF"],
    "pdf":   ["Tóm tắt", "Trích xuất dữ liệu", "Chuyển sang Word"],
    "image": ["Trích xuất dữ liệu", "Chuyển sang PDF"],
    None:    ["Tóm tắt", "Soạn công văn", "Tính lương"],
}

# Định dạng đích cho modal Chuyển đổi định dạng
CONVERT_TARGETS = {
    "word":  ["PDF (.pdf)"],
    "excel": ["PDF (.pdf)"],
    "ppt":   ["PDF (.pdf)"],
    "pdf":   ["Word (.docx)", "Excel (.xlsx)", "Ảnh (.png)"],
    "image": ["PDF (.pdf)"],
}
CONVERT_EXT = {
    "PDF (.pdf)": ".pdf", "Word (.docx)": ".docx",
    "Excel (.xlsx)": ".xlsx", "Ảnh (.png)": ".png",
}
