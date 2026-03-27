LIGHT_STYLESHEET = """
QWidget {
    background: #f5f7fb;
    color: #1f2937;
    font-family: "Segoe UI", "PingFang SC", "Microsoft YaHei";
    font-size: 13px;
}
QMainWindow, QFrame#Card, QFrame#HeroCard, QGroupBox {
    background: #ffffff;
}
QFrame#HeroCard {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #eff6ff, stop:1 #f8fafc);
    border: 1px solid #dbeafe;
    border-radius: 16px;
}
QFrame#Card, QGroupBox {
    border: 1px solid #e5e7eb;
    border-radius: 12px;
}
QLabel[role="pageTitle"] {
    font-size: 22px;
    font-weight: 700;
    color: #0f172a;
}
QLabel[role="pageSubtitle"] {
    font-size: 13px;
    color: #475569;
}
QLabel[role="sectionTitle"] {
    font-size: 16px;
    font-weight: 700;
    color: #0f172a;
}
QLabel[role="muted"] {
    color: #64748b;
}
QLabel[role="hint"] {
    color: #475569;
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 10px 12px;
}
QLabel[role="good"] {
    color: #166534;
    background: #f0fdf4;
    border: 1px solid #bbf7d0;
    border-radius: 10px;
    padding: 10px 12px;
}
QLabel[role="warn"] {
    color: #92400e;
    background: #fffbeb;
    border: 1px solid #fde68a;
    border-radius: 10px;
    padding: 10px 12px;
}
QLabel[role="statTitle"] {
    color: #64748b;
    font-size: 12px;
}
QLabel[role="statValue"] {
    color: #0f172a;
    font-size: 20px;
    font-weight: 700;
}
QGroupBox {
    margin-top: 12px;
    padding-top: 14px;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 12px;
    padding: 0 4px;
}
QPushButton {
    background: #2563eb;
    color: #ffffff;
    border: none;
    border-radius: 8px;
    padding: 8px 14px;
}
QPushButton:hover {
    background: #1d4ed8;
}
QPushButton[variant="secondary"] {
    background: #e5eefc;
    color: #1d4ed8;
}
QPushButton[variant="ghost"] {
    background: #ffffff;
    color: #334155;
    border: 1px solid #cbd5e1;
}
QPushButton[variant="danger"] {
    background: #dc2626;
}
QLineEdit, QTextEdit, QPlainTextEdit, QTableWidget, QListWidget, QSpinBox, QDoubleSpinBox, QComboBox {
    background: #ffffff;
    border: 1px solid #d1d5db;
    border-radius: 8px;
    padding: 6px;
}
QHeaderView::section {
    background: #eff6ff;
    color: #1e3a8a;
    border: none;
    border-bottom: 1px solid #dbeafe;
    padding: 8px;
    font-weight: 600;
}
QTableWidget {
    gridline-color: #e5e7eb;
}
QListWidget::item {
    border-radius: 8px;
    padding: 10px 12px;
    margin: 2px 6px;
}
QListWidget::item:selected {
    background: #dbeafe;
    color: #1d4ed8;
}
QListWidget#SideNav {
    border: none;
    background: transparent;
}
QListWidget#SideNav::item {
    margin: 4px 0;
    padding: 12px 14px;
}
QListWidget#SideNav::item:selected {
    background: #ffffff;
    border: 1px solid #bfdbfe;
}
QStatusBar {
    background: #ffffff;
}
"""
