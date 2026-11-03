"""Application-wide Qt stylesheet."""

from __future__ import annotations

COLORS = {
    "ink": "#162033",
    "muted": "#687A91",
    "faint": "#94A3B8",
    "primary": "#2563EB",
    "primary_dark": "#1D4ED8",
    "sidebar": "#0B1729",
    "success": "#0E8A61",
    "danger": "#C74040",
}

APP_QSS = r"""
QWidget {
    color: #182235;
    font-family: "Microsoft YaHei UI", "Microsoft YaHei", "Segoe UI";
    font-size: 13px;
}
QMainWindow, QWidget#Root {
    background: #F5F7FB;
}
QToolTip {
    background: #17243A;
    color: #F8FAFC;
    border: 1px solid #2A3A53;
    border-radius: 5px;
    padding: 5px 7px;
    font-size: 11px;
}

/* ---------- Sidebar ---------- */
QFrame#Sidebar {
    background: #0B1729;
    border: none;
}
QLabel#BrandTitle {
    color: #FFFFFF;
    font-size: 20px;
    font-weight: 700;
}
QLabel#BrandSub {
    color: #6E86A5;
    font-size: 9px;
    font-weight: 600;
    letter-spacing: 2px;
}
QLabel#SidebarSection {
    color: #5E7795;
    font-size: 10px;
    font-weight: 600;
    letter-spacing: 1px;
}
QPushButton#NavButton {
    color: #92A6C0;
    background: transparent;
    border: none;
    border-radius: 10px;
    text-align: left;
    padding: 0 14px;
    font-size: 14px;
    font-weight: 500;
}
QPushButton#NavButton:hover {
    background: #12233B;
    color: #DCE8F7;
}
QPushButton#NavButton:checked {
    background: #1A3353;
    color: #FFFFFF;
    font-weight: 600;
}
QFrame#EngineCard {
    background: #112239;
    border: 1px solid #1C3350;
    border-radius: 12px;
}
QLabel#EngineTitle {
    color: #728BA8;
    font-size: 10px;
    font-weight: 600;
    letter-spacing: 1px;
}
QLabel#EngineLine {
    color: #BAC9DB;
    font-size: 11px;
}
QLabel#EngineDotOn {
    color: #43D6A0;
    font-size: 10px;
}
QLabel#EngineDotOff {
    color: #F18B82;
    font-size: 10px;
}
QLabel#PrivacyLine {
    color: #6B829F;
    font-size: 10px;
}
QLabel#VersionLabel {
    color: #49637F;
    font-size: 9px;
}

/* ---------- Page header ---------- */
QLabel#PageTitle {
    color: #101B2E;
    font-size: 26px;
    font-weight: 700;
}
QLabel#PageSubtitle {
    color: #718299;
    font-size: 12px;
}
QLabel#SectionTitle {
    color: #1B283B;
    font-size: 14px;
    font-weight: 600;
}
QLabel#CountBadge {
    background: #EAF0FB;
    color: #2D5BB9;
    border-radius: 8px;
    padding: 1px 7px;
    font-size: 10px;
    font-weight: 700;
}
QLabel#QueueHint {
    color: #8A99AC;
    font-size: 11px;
}

/* ---------- Drop zone ---------- */
QFrame#DropArea {
    background: #FFFFFF;
    border: 1px dashed #C2D2E6;
    border-radius: 14px;
}
QFrame#DropArea[dragActive="true"] {
    background: #F1F7FF;
    border: 1px dashed #2563EB;
}
QLabel#DropIcon {
    background: #EDF4FF;
    border-radius: 21px;
}
QLabel#DropTitle {
    color: #1A2940;
    font-size: 14px;
    font-weight: 600;
}
QLabel#DropHint {
    color: #7A8CA4;
    font-size: 11px;
}

/* ---------- Cards / queue ---------- */
QFrame#Card, QFrame#QueueCard {
    background: #FFFFFF;
    border: 1px solid #E6ECF4;
    border-radius: 14px;
}
QLabel#EmptyTitle {
    color: #53667E;
    font-size: 13px;
    font-weight: 500;
}
QLabel#EmptyHint {
    color: #96A4B7;
    font-size: 11px;
}
QTableWidget {
    background: #FFFFFF;
    border: none;
    gridline-color: transparent;
    selection-background-color: #F4F8FF;
    selection-color: #182235;
    alternate-background-color: #FFFFFF;
    outline: none;
}
QTableWidget::item {
    border-bottom: 1px solid #F0F3F8;
    padding: 0;
}
QTableWidget::item:selected {
    background: #F4F8FF;
    color: #182235;
}
QHeaderView::section {
    background: #FFFFFF;
    color: #8A98AB;
    border: none;
    border-bottom: 1px solid #EDF1F7;
    padding: 7px 10px;
    font-size: 10px;
    font-weight: 600;
}
QTableCornerButton::section {
    background: #FFFFFF;
    border: none;
}
QLabel#FileBadgeWord {
    background: #EAF1FF;
    color: #2356C9;
    border-radius: 8px;
    font-size: 14px;
    font-weight: 700;
}
QLabel#FileBadgePdf {
    background: #FFF0EE;
    color: #D4483F;
    border-radius: 8px;
    font-size: 9px;
    font-weight: 700;
}
QLabel#FileName {
    color: #1B283B;
    font-size: 12px;
    font-weight: 600;
}
QLabel#FilePath {
    color: #97A5B8;
    font-size: 10px;
}
QLabel#FileSize {
    color: #5F7188;
    font-size: 11px;
}
QLabel#StatusPill {
    background: #F1F4F8;
    color: #66788F;
    border-radius: 9px;
    padding: 2px 7px;
    font-size: 10px;
    font-weight: 600;
}
QLabel#StatusPill[state="working"] {
    background: #EAF1FF;
    color: #2356C9;
}
QLabel#StatusPill[state="success"] {
    background: #E8F7F1;
    color: #0C825B;
}
QLabel#StatusPill[state="error"] {
    background: #FDECEC;
    color: #C03939;
}
QLabel#StatusPill[state="cancel"] {
    background: #F3F4F6;
    color: #7A8798;
}
QToolButton#RowAction {
    background: transparent;
    border: none;
    border-radius: 6px;
}
QToolButton#RowAction:hover {
    background: #F3F6FB;
}
QScrollBar:vertical {
    background: transparent;
    width: 9px;
    margin: 2px 1px;
}
QScrollBar::handle:vertical {
    background: #D5DEEA;
    border-radius: 4px;
    min-height: 28px;
}
QScrollBar::handle:vertical:hover {
    background: #BCC9DA;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0;
}
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
    background: transparent;
}
QScrollBar:horizontal {
    height: 0;
}

/* ---------- Buttons ---------- */
QPushButton#PrimaryButton {
    background: #2563EB;
    color: #FFFFFF;
    border: none;
    border-radius: 10px;
    padding: 0 22px;
    font-size: 13px;
    font-weight: 600;
}
QPushButton#PrimaryButton:hover {
    background: #1E54D4;
}
QPushButton#PrimaryButton:pressed {
    background: #1948BC;
}
QPushButton#PrimaryButton:disabled {
    background: #B9CAE7;
    color: #F6F9FF;
}
QPushButton#SecondaryButton {
    background: #FFFFFF;
    color: #31445D;
    border: 1px solid #D7E0EC;
    border-radius: 9px;
    padding: 0 14px;
    font-size: 12px;
    font-weight: 500;
}
QPushButton#SecondaryButton:hover {
    border-color: #AFC2DE;
    background: #F8FAFE;
}
QPushButton#SecondaryButton:pressed {
    background: #F0F4FA;
}
QPushButton#SecondaryButton:disabled {
    color: #A8B4C3;
    border-color: #E6EBF2;
}
QPushButton#GhostButton {
    background: transparent;
    color: #5D7089;
    border: none;
    border-radius: 7px;
    padding: 0 8px;
    font-size: 11px;
    font-weight: 500;
}
QPushButton#GhostButton:hover {
    color: #2356C9;
    background: #EDF3FC;
}
QPushButton#GhostButton:disabled {
    color: #B4BFCD;
}
QPushButton#DangerGhost {
    background: transparent;
    color: #C34848;
    border: 1px solid #F0D2D2;
    border-radius: 9px;
    padding: 0 15px;
    font-size: 12px;
    font-weight: 500;
}
QPushButton#DangerGhost:hover {
    background: #FFF6F6;
    border-color: #E7B9B9;
}
QPushButton#LinkButton {
    background: transparent;
    color: #2360D6;
    border: none;
    padding: 0 3px;
    font-size: 11px;
    font-weight: 600;
}
QPushButton#LinkButton:hover {
    color: #1747A8;
    text-decoration: underline;
}

/* ---------- Output / footer ---------- */
QLabel#FieldLabel {
    color: #5A6E86;
    font-size: 11px;
    font-weight: 600;
}
QRadioButton {
    color: #35465C;
    font-size: 11px;
    spacing: 7px;
}
QRadioButton::indicator {
    width: 11px;
    height: 11px;
    border-radius: 7px;
    border: 1px solid #B7C5D8;
    background: #FFFFFF;
}
QRadioButton::indicator:hover {
    border-color: #7295C9;
}
QRadioButton::indicator:checked {
    border: 3px solid #2563EB;
    background: #FFFFFF;
}
QLineEdit#OutputPath {
    background: #F7F9FC;
    border: 1px solid #E4EAF3;
    border-radius: 8px;
    padding: 5px 9px;
    color: #5D7087;
    font-size: 11px;
}
QLineEdit#OutputPath:disabled {
    color: #A2AEBD;
    background: #F8FAFC;
}
QLabel#StatusText {
    color: #677A91;
    font-size: 11px;
}
QLabel#StatusText[state="success"] {
    color: #0C825B;
}
QLabel#StatusText[state="error"] {
    color: #C03939;
}
QLabel#StatusText[state="working"] {
    color: #2356C9;
}
QProgressBar {
    background: #EAF0F8;
    border: none;
    border-radius: 3px;
    min-height: 6px;
    max-height: 6px;
    text-align: center;
    color: transparent;
}
QProgressBar::chunk {
    background: #2E68E8;
    border-radius: 3px;
}
QLabel#FooterHint {
    color: #95A3B6;
    font-size: 10px;
}
QFrame#Divider {
    background: #EEF2F7;
    max-height: 1px;
    border: none;
}

/* ---------- Dialogs ---------- */
QDialog {
    background: #FFFFFF;
}
QLabel#DialogTitle {
    color: #152135;
    font-size: 18px;
    font-weight: 700;
}
QLabel#DialogSubtitle {
    color: #718299;
    font-size: 12px;
}
QTextEdit#DetailsView {
    background: #F7F9FC;
    border: 1px solid #E3E9F2;
    border-radius: 9px;
    color: #40536B;
    font-family: "Cascadia Mono", "Consolas", "Microsoft YaHei UI";
    font-size: 11px;
    padding: 10px;
}
QMessageBox {
    background: #FFFFFF;
}
QMessageBox QLabel {
    color: #25364B;
    font-size: 12px;
}
"""
