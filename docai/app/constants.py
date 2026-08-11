APP_NAME = "DocAI"
APP_TAGLINE = "Trợ lý AI xử lý tài liệu văn phòng"

# Plan & quota (mock — backend Phase 2)
PLAN_NAME = "Pro"
PLAN_QUOTA = 500

# File type detection
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

# Suggestion chips on the central AI chat (label · icon ◆ color) — per design
EMPTY_CHIPS = [
    ("Tóm tắt hợp đồng dài",       "#C0392B"),
    ("Trích xuất hóa đơn → Excel", "#1F8A5B"),
    ("Tính lương + BHXH",          "#2A6FDB"),
    ("Soạn công văn hành chính",   "#6E6A63"),
]

# File-type badge for a "Recent" item in the sidebar: (label, color key)
FILE_BADGE_STYLE = {
    "word":  ("DOC", "doc"),
    "excel": ("XLS", "xls"),
    "pdf":   ("PDF", "pdf"),
    "ppt":   ("PPT", "other"),
    "image": ("IMG", "other"),
    "xml":   ("XML", "other"),
}

# Folder mode: also recognize XML (e-invoices) — only used to check it in
# as context, doesn't need a preview renderer like File mode does.
FOLDER_EXT_MAP = {**EXT_MAP, ".xml": "xml"}
FOLDER_BADGE = {**FILE_BADGE, "xml": "XML"}

# Quick-action chips based on the currently open file's context
CONTEXT_CHIPS = {
    "word":  ["Tóm tắt", "Soạn công văn", "Chuyển sang PDF"],
    "excel": ["Tính lương", "Tạo bảng kê thuế", "Chuyển sang PDF"],
    "ppt":   ["Tóm tắt", "Chuyển sang PDF"],
    "pdf":   ["Tóm tắt", "Trích xuất dữ liệu", "Chuyển sang Word"],
    "image": ["Trích xuất dữ liệu", "Chuyển sang PDF"],
    None:    ["Tóm tắt", "Soạn công văn", "Tính lương"],
}

# Create a new document via chat (no attachment needed beforehand) — type
# picker card at 7.2: (type key, label, badge, text color, badge background)
CREATE_TYPE_OPTIONS = [
    ("word",  "Word",       "DOC", "#2A6FDB", "#EAF1FC"),
    ("excel", "Excel",      "XLS", "#1F8A5B", "#E6F4EC"),
    ("ppt",   "PowerPoint", "PPT", "#C0392B", "#FBECEA"),
]
CREATE_EXT = {"word": ".docx", "excel": ".xlsx", "ppt": ".pptx"}
CREATE_TITLE = {
    "word": "Tạo văn bản mới", "excel": "Tạo bảng tính mới", "ppt": "Tạo bản trình bày mới",
}

# App label for the "Open with …" button in the preview panel (7.6)
OPEN_WITH_LABEL = {
    "word": "Word", "excel": "Excel", "ppt": "PowerPoint",
    "pdf": "PDF", "image": "ảnh",
}

# Target formats for the Convert Format modal
CONVERT_TARGETS = {
    "word":  ["PDF (.pdf)"],
    "excel": ["PDF (.pdf)"],
    "ppt":   ["PDF (.pdf)"],
    "pdf":   ["Word (.docx)", "Excel (.xlsx)", "Ảnh (.png)"],
    "image": [
        "PDF (.pdf)", "PDF - có thể tìm kiếm (.pdf)", "Word (.docx)", "Excel (.xlsx)",
        "PNG (.png)", "JPG (.jpg)", "WEBP (.webp)", "BMP (.bmp)",
    ],
}
CONVERT_EXT = {
    "PDF (.pdf)": ".pdf", "PDF - có thể tìm kiếm (.pdf)": ".pdf", "Word (.docx)": ".docx",
    "Excel (.xlsx)": ".xlsx", "Ảnh (.png)": ".png",
    "PNG (.png)": ".png", "JPG (.jpg)": ".jpg",
    "WEBP (.webp)": ".webp", "BMP (.bmp)": ".bmp",
}
