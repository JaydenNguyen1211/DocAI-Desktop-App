# Theme sáng theo wireframe: nền kem, thẻ trắng, nút chính đen, accent xanh.

COLORS = {
    "bg":        "#E4E3DC",   # nền ngoài (kem)
    "surface":   "#FFFFFF",   # thẻ / panel trắng
    "sidebar":   "#F2F1EC",
    "border":    "#C9C8C0",
    "border_soft": "#DDDCD4",
    "text":      "#1A1A1A",
    "muted":     "#8A8A85",
    "primary":   "#1A1A1A",   # nút đen
    "accent":    "#2563EB",   # xanh link / selected
    "accent_bg": "#E8F0FE",   # bong bóng user
    "ai_bg":     "#F1F0EB",   # bong bóng AI
    "danger":    "#DC2626",
    "danger_bg": "#FEF2F2",
}

APP_STYLE = f"""
* {{ font-family: 'Segoe UI', 'Arial', sans-serif; }}

QMainWindow, QDialog {{
    background-color: {COLORS['bg']};
}}
QWidget {{
    color: {COLORS['text']};
    font-size: 13px;
}}

/* ── Top bar ─────────────────────────────────────────────── */
QFrame#topbar {{
    background-color: {COLORS['surface']};
    border-bottom: 1px solid {COLORS['border']};
}}
QLabel#appLogoTile {{
    background: {COLORS['primary']}; color: white;
    border-radius: 7px; font-size: 13px; font-weight: 700;
}}
QLabel#appTitle {{
    font-size: 15px; font-weight: 600;
}}
QLabel#quotaLabel {{
    color: {COLORS['muted']}; font-size: 12px;
}}
QPushButton#gearBtn {{
    background: transparent; border: none; font-size: 17px;
    color: {COLORS['muted']}; padding: 4px 8px; border-radius: 6px;
}}
QPushButton#gearBtn:hover {{ background: {COLORS['sidebar']}; color: {COLORS['text']}; }}

/* ── Sidebar ─────────────────────────────────────────────── */
QFrame#sidebar {{
    background-color: {COLORS['sidebar']};
    border-right: 1px solid {COLORS['border']};
}}
QFrame#segControl {{
    background: #E7E3DC;
    border-radius: 9px;
}}
QPushButton.segBtn {{
    background: transparent; border: none;
    border-radius: 6px; padding: 6px 0;
    color: {COLORS['text']}; font-size: 12.5px; font-weight: 500;
}}
QPushButton.segBtn:checked {{
    background: {COLORS['surface']};
    color: {COLORS['text']}; font-weight: 700;
}}

QPushButton#newFileBtn {{
    background: {COLORS['surface']};
    border: 1px solid {COLORS['border']};
    border-radius: 8px; padding: 9px 12px;
    color: {COLORS['text']}; font-weight: 700; text-align: center;
}}
QPushButton#newFileBtn:hover {{ border-color: {COLORS['text']}; }}
QLabel#recentLabel {{
    color: {COLORS['muted']}; font-size: 11px; font-weight: 600;
    padding-left: 4px; margin-top: 2px;
}}
QLabel#recentEmpty {{
    color: {COLORS['muted']}; font-size: 12px; padding-left: 6px;
}}

/* Item lịch sử — thẻ 2 hàng: tiêu đề + (badge · tên file).
   Nền trắng khi active/hover được vẽ bằng QPainter trong RecentItem.paintEvent. */
QLabel#recentTitle {{
    background: transparent; font-size: 12.5px; font-weight: 600; color: #3A3833;
}}
QLabel#recentMeta {{ background: transparent; font-size: 10.5px; color: #A8A49C; }}
QLabel.fileBadgeDoc, QLabel.fileBadgePdf,
QLabel.fileBadgeXls, QLabel.fileBadgeOther {{
    font-size: 8px; font-weight: 700; border-radius: 4px; padding: 2px 5px;
}}
QLabel.fileBadgeDoc   {{ color: #2A6FDB; background: #EAF1FC; }}
QLabel.fileBadgePdf   {{ color: #C0392B; background: #FBECEA; }}
QLabel.fileBadgeXls   {{ color: #1F8A5B; background: #E6F4EC; }}
QLabel.fileBadgeOther {{ color: #6E6A63; background: #EFEDE9; }}

/* Thẻ gói dịch vụ + thanh tiến trình */
QFrame#planCard {{
    background: {COLORS['surface']};
    border: 1px solid {COLORS['border_soft']};
    border-radius: 11px;
}}
QLabel#planName {{ background: transparent; font-size: 12.5px; font-weight: 600; }}
QLabel#planCount {{ background: transparent; font-size: 11px; color: {COLORS['muted']}; }}
QLabel#planHint {{ background: transparent; font-size: 10.5px; color: {COLORS['muted']}; }}
QProgressBar#planBar {{
    background: #EFEDE9; border: none; border-radius: 3px;
}}
QProgressBar#planBar::chunk {{
    background: {COLORS['accent']}; border-radius: 3px;
}}

/* ── Chat AI trung tâm (Trạng thái 1) ────────────────────── */
QFrame#centralChat {{ background: {COLORS['surface']}; }}
QLabel#welcomeLogo {{
    background: {COLORS['accent_bg']};
    border: 1px solid #D4E3F8;
    border-radius: 14px;
    color: {COLORS['accent']};
    font-size: 26px; font-weight: 700;
}}
QLabel#welcomeTitle {{
    font-size: 19px; font-weight: 700; letter-spacing: -0.2px;
}}
QLabel#welcomeSub {{ color: {COLORS['muted']}; font-size: 12.5px; }}

QFrame#welcomeInput {{
    background: {COLORS['surface']};
    border: 1px solid {COLORS['border']};
    border-radius: 13px;
}}
QFrame#welcomeInput[dragOver="true"] {{
    border-color: {COLORS['accent']};
    background: {COLORS['accent_bg']};
}}
QLineEdit#welcomeInputBox {{
    background: transparent; border: none;
    padding: 0 2px; font-size: 13px;
}}
QLineEdit#welcomeInputBox:focus {{ border: none; }}
QPushButton#welcomeSendBtn {{
    background: {COLORS['primary']}; color: white; border: none;
    border-radius: 8px;
}}
QPushButton#welcomeSendBtn:hover {{ background: #333333; }}
QPushButton#welcomeSendBtn:disabled {{ background: #9B9B96; }}

QPushButton#attachBtn {{
    background: transparent;
    border: 1px solid {COLORS['border_soft']};
    border-radius: 8px;
}}
QPushButton#attachBtn:hover {{ border-color: {COLORS['accent']}; }}

QPushButton.chip {{
    background: {COLORS['surface']};
    border: 1px solid {COLORS['border']};
    border-radius: 15px; padding: 7px 16px; font-size: 13px;
}}
QPushButton.chip:hover {{ border-color: {COLORS['accent']}; color: {COLORS['accent']}; }}

/* ── Preview panel ───────────────────────────────────────── */
QFrame#previewPanel {{
    background: {COLORS['surface']};
    border: 1px solid {COLORS['border']};
    border-radius: 10px;
}}
QLabel#previewTitle {{ font-size: 13px; font-weight: 600; }}
QLabel#previewBadge {{
    color: {COLORS['muted']}; font-size: 11px; font-weight: 700;
    border: 1px solid {COLORS['border']}; border-radius: 4px;
    padding: 1px 6px;
}}
QLabel#previewFooter {{
    color: {COLORS['muted']}; font-size: 11px;
    border-top: 1px solid {COLORS['border_soft']}; padding: 6px;
}}
QPushButton#convertBtn {{
    background: transparent; border: 1px solid {COLORS['border']};
    border-radius: 6px; padding: 4px 12px; font-size: 12px;
}}
QPushButton#convertBtn:hover {{ border-color: {COLORS['text']}; }}

/* ── Chat panel ──────────────────────────────────────────── */
QFrame#chatPanel {{
    background: {COLORS['surface']};
    border: 1px solid {COLORS['border']};
    border-radius: 10px;
}}
QLabel#chatHeader {{ color: {COLORS['muted']}; font-size: 13px; }}
QScrollArea#chatScroll {{ background: transparent; border: none; }}
QWidget#chatMsgHost {{ background: transparent; }}
QWidget#chatChipsHost {{ background: transparent; }}

QLabel.bubbleUser {{
    background: #EAF1FC;
    color: #173A66;
    border-radius: 10px; padding: 10px 14px; font-size: 13px;
}}
QLabel.bubbleAI {{
    background: #F5F4F1;
    border: 1px solid #EAE8E3;
    border-radius: 10px; padding: 10px 14px; font-size: 13px;
}}
QLabel.bubbleSystem {{
    color: {COLORS['muted']}; font-size: 12px; padding: 2px 6px;
}}
QPushButton.fileCard {{
    background: {COLORS['ai_bg']};
    border: 1px solid {COLORS['border_soft']};
    border-radius: 10px; padding: 10px 14px;
    color: {COLORS['accent']}; font-size: 13px; text-align: left;
}}
QPushButton.fileCard:hover {{ border-color: {COLORS['accent']}; }}

QFrame#inputArea {{
    background: transparent;
    border-top: 1px solid {COLORS['border_soft']};
}}
QLineEdit#inputBox {{
    background: {COLORS['ai_bg']};
    border: 1px solid {COLORS['border_soft']};
    border-radius: 10px; padding: 10px 14px; font-size: 13px;
}}
QLineEdit#inputBox:focus {{ border-color: {COLORS['accent']}; }}
QPushButton#sendBtn {{
    background: {COLORS['primary']}; color: white; border: none;
    border-radius: 8px;
}}
QPushButton#sendBtn:hover {{ background: #333333; }}
QPushButton#sendBtn:disabled {{ background: #9B9B96; }}

/* ── Splash ──────────────────────────────────────────────── */
QFrame#splashCard {{
    background: {COLORS['surface']};
    border: 1px solid {COLORS['text']};
    border-radius: 16px;
}}
QLabel#splashLogo {{
    background: {COLORS['accent_bg']};
    border: 2px solid {COLORS['accent']};
    border-radius: 10px;
    color: {COLORS['accent']};
    font-size: 24px; font-weight: 700;
}}
QLabel#splashTitle {{ font-size: 24px; font-weight: 700; }}
QLabel#splashSub {{ color: {COLORS['muted']}; font-size: 13px; }}
QLabel#splashStatus {{ color: {COLORS['muted']}; font-size: 12px; }}
QLabel#splashError {{ color: {COLORS['danger']}; font-size: 12px; }}

/* ── Modals / shared ─────────────────────────────────────── */
QLabel#modalTitle {{ font-size: 18px; font-weight: 700; }}
QLabel#fieldLabel {{ color: {COLORS['muted']}; font-size: 13px; }}
QLabel#valueLabel {{ font-size: 13px; font-weight: 700; }}
QLabel#footnote {{ color: {COLORS['muted']}; font-size: 11px; }}
QLabel#dangerIcon {{
    background: {COLORS['danger_bg']};
    border: 2px solid {COLORS['danger']};
    border-radius: 8px; color: {COLORS['danger']};
    font-size: 15px; font-weight: 700;
}}
QPushButton#modalClose {{
    background: transparent; border: none;
    color: {COLORS['muted']}; font-size: 15px; padding: 2px 8px;
}}
QPushButton#modalClose:hover {{ color: {COLORS['text']}; }}

QPushButton.tabBtn {{
    background: {COLORS['surface']};
    border: 1px solid {COLORS['border']};
    border-radius: 7px; padding: 7px 18px; font-size: 13px;
    color: {COLORS['muted']};
}}
QPushButton.tabBtn:checked {{
    background: {COLORS['accent_bg']};
    border-color: {COLORS['accent']};
    color: {COLORS['accent']}; font-weight: 600;
}}

QPushButton.primaryBtn {{
    background: {COLORS['primary']};
    color: white; border: none;
    border-radius: 8px; padding: 10px 22px;
    font-weight: 700; font-size: 13px;
}}
QPushButton.primaryBtn:hover {{ background: #333333; }}
QPushButton.primaryBtn:disabled {{ background: #9B9B96; }}

QPushButton.secondaryBtn {{
    background: {COLORS['surface']};
    border: 1px solid {COLORS['border']};
    border-radius: 8px; padding: 10px 22px;
    font-weight: 600; font-size: 13px;
}}
QPushButton.secondaryBtn:hover {{ border-color: {COLORS['text']}; }}

QPushButton.linkBtn {{
    background: transparent;
    border: none;
    color: {COLORS['text']};
    font-weight: 700; font-size: 13px;
    padding: 2px 4px;
}}
QPushButton.linkBtn:hover {{ text-decoration: underline; }}
QPushButton.linkBtn:disabled {{ color: #9B9B96; }}

QLineEdit {{
    background: {COLORS['surface']};
    border: 1px solid {COLORS['border']};
    border-radius: 8px; padding: 9px 12px; font-size: 13px;
}}
QLineEdit:focus {{ border-color: {COLORS['accent']}; }}
QLineEdit[readOnly="true"] {{ background: {COLORS['sidebar']}; }}

QComboBox {{
    background: {COLORS['sidebar']};
    border: 1px solid {COLORS['border']};
    border-radius: 8px; padding: 8px 12px; font-size: 13px;
}}
QComboBox::drop-down {{ border: none; width: 24px; }}
QComboBox QAbstractItemView {{
    background: {COLORS['surface']};
    border: 1px solid {COLORS['border']};
    selection-background-color: {COLORS['accent_bg']};
    selection-color: {COLORS['text']};
    outline: none;
}}

QProgressBar {{
    background: {COLORS['border_soft']};
    border: none; border-radius: 4px; height: 8px;
    text-align: center; color: transparent;
}}
QProgressBar::chunk {{
    background: {COLORS['accent']}; border-radius: 4px;
}}

QScrollBar:vertical {{
    background: transparent; width: 8px; margin: 0;
}}
QScrollBar::handle:vertical {{
    background: {COLORS['border']}; border-radius: 4px; min-height: 24px;
}}
QScrollBar::handle:vertical:hover {{ background: {COLORS['muted']}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{ background: none; }}

QSplitter::handle:horizontal {{ background: transparent; width: 10px; }}

/* ── Toast prompt-caching ────────────────────────────────── */
QLabel#cacheToast {{
    border-radius: 8px; padding: 8px 14px; font-size: 11.5px; font-weight: 600;
}}
QLabel.cacheToastHit {{
    background: #E6F4EC; color: #1F8A5B; border: 1px solid #C7E8D6;
}}
QLabel.cacheToastNew {{
    background: {COLORS['accent_bg']}; color: {COLORS['accent']}; border: 1px solid #D4E3F8;
}}
QLabel.cacheToastNone {{
    background: {COLORS['sidebar']}; color: {COLORS['muted']}; border: 1px solid {COLORS['border']};
}}
"""
