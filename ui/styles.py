"""
styles.py — Общие стили приложения (современная светлая тема).
Используются в MainWindow и в модальных диалогах (SettingsDialog и др.).
"""

STYLESHEET = """
/* ─── Современная светлая тема ─── */
QMainWindow, QDialog {
    background-color: #f1f5f9;
}
QWidget#centralWidget {
    background-color: #f1f5f9;
}
/* ─── Заголовок ─── */
QWidget#headerBar {
    background: qlineargradient(x1:0 y1:0 x2:0 y2:1, stop:0 #0f172a, stop:1 #1e293b);
    min-height: 52px;
    max-height: 52px;
    border-bottom: 1px solid rgba(148, 163, 184, 0.2);
}
QLabel#appTitle {
    color: #f8fafc;
    font-size: 17px;
    font-weight: 600;
    letter-spacing: 0.4px;
    padding-left: 24px;
}
QLabel#pageTitle {
    color: #94a3b8;
    font-size: 13px;
    padding-right: 20px;
}
/* ─── Меню ─── */
QMenuBar {
    background: transparent;
    color: #cbd5e1;
    font-size: 13px;
    spacing: 4px;
    padding: 0 12px;
}
QMenuBar::item {
    background: transparent;
    padding: 8px 14px;
    border-radius: 6px;
}
QMenuBar::item:selected {
    background-color: rgba(255,255,255,0.12);
    color: #f8fafc;
}
QMenu {
    background-color: #1e293b;
    color: #e2e8f0;
    border: 1px solid #475569;
    border-radius: 8px;
    padding: 6px 0;
    font-size: 13px;
}
QMenu::item {
    padding: 10px 24px;
}
QMenu::item:selected {
    background-color: #334155;
    color: #f8fafc;
}
QMenu::separator {
    height: 1px;
    background: #475569;
    margin: 6px 12px;
}
/* ─── Кнопки ─── */
QPushButton {
    font-size: 13px;
    border-radius: 8px;
    padding: 10px 20px;
    font-weight: 500;
}
QPushButton#btnPrimary {
    background: qlineargradient(x1:0 y1:0 x2:0 y2:1, stop:0 #3b82f6, stop:1 #2563eb);
    color: white;
    border: none;
}
QPushButton#btnPrimary:hover {
    background: qlineargradient(x1:0 y1:0 x2:0 y2:1, stop:0 #2563eb, stop:1 #1d4ed8);
}
QPushButton#btnPrimary:pressed {
    background: #1d4ed8;
}
QPushButton#btnPrimary:disabled {
    background: #94a3b8;
    color: #e2e8f0;
}
QPushButton#btnSecondary {
    background-color: #ffffff;
    color: #334155;
    border: 1px solid #e2e8f0;
}
QPushButton#btnSecondary:hover {
    background-color: #f1f5f9;
    border-color: #cbd5e1;
}
QPushButton#btnDanger {
    background-color: #fef2f2;
    color: #dc2626;
    border: 1px solid #fecaca;
}
QPushButton#btnDanger:hover {
    background-color: #fee2e2;
}
QPushButton#btnBack {
    background-color: transparent;
    color: #64748b;
    border: 1px solid #e2e8f0;
    padding: 8px 16px;
}
QPushButton#btnBack:hover {
    background-color: #f1f5f9;
    color: #334155;
}
QPushButton#btnIcon {
    background-color: transparent;
    color: #64748b;
    border: none;
    padding: 6px 10px;
    font-size: 16px;
    border-radius: 6px;
}
QPushButton#btnIcon:hover {
    background-color: rgba(0,0,0,0.06);
    color: #1e293b;
}
QPushButton#btnIconDanger {
    background-color: transparent;
    color: #dc2626;
    border: none;
    padding: 6px 10px;
    font-size: 16px;
    border-radius: 6px;
}
QPushButton#btnIconDanger:hover {
    background-color: #fef2f2;
    color: #b91c1c;
}
/* ─── Разделитель ─── */
QFrame#separator {
    background-color: #e2e8f0;
    max-height: 1px;
}
/* ─── Карточки ─── */
QFrame#card {
    background-color: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 4px;
}
/* ─── Вкладки ─── */
QTabWidget::pane {
    background-color: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    margin-top: -1px;
    padding: 12px;
}
QTabBar::tab {
    background-color: #f8fafc;
    color: #64748b;
    padding: 10px 20px;
    margin-right: 4px;
    border: 1px solid #e2e8f0;
    border-bottom: none;
    border-top-left-radius: 8px;
    border-top-right-radius: 8px;
    font-weight: 500;
}
QTabBar::tab:selected {
    background-color: #ffffff;
    color: #1e293b;
    border-color: #e2e8f0;
    border-bottom: 1px solid #ffffff;
}
QTabBar::tab:hover:!selected {
    background-color: #f1f5f9;
    color: #475569;
}
/* ─── Группы ─── */
QGroupBox {
    font-size: 13px;
    font-weight: 600;
    color: #334155;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    margin-top: 12px;
    padding: 16px 16px 8px 16px;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 14px;
    padding: 0 8px;
    background-color: #ffffff;
    color: #475569;
}
/* ─── Поля ввода ─── */
QLineEdit, QTextEdit {
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    padding: 8px 12px;
    font-size: 13px;
    background: #ffffff;
    color: #1e293b;
    selection-background-color: #dbeafe;
}
QLineEdit:focus, QTextEdit:focus {
    border-color: #3b82f6;
    outline: none;
}
/* ─── Таблицы ─── */
QTableWidget, QTableView {
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    background: #ffffff;
    gridline-color: #f1f5f9;
    font-size: 13px;
    color: #1e293b;
}
QTableWidget::item, QTableView {
    padding: 8px 10px;
}
QTableWidget::item:selected, QTableView::item:selected {
    background-color: #eff6ff;
    color: #1e293b;
}
QHeaderView::section {
    background-color: #f8fafc;
    color: #475569;
    font-weight: 600;
    font-size: 12px;
    padding: 10px 12px;
    border: none;
    border-bottom: 2px solid #e2e8f0;
    border-right: 1px solid #f1f5f9;
}
/* ─── ComboBox ─── */
QComboBox {
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    padding: 8px 12px;
    font-size: 13px;
    background: #ffffff;
    color: #1e293b;
    min-width: 160px;
}
QComboBox:focus {
    border-color: #3b82f6;
}
QComboBox::drop-down {
    border: none;
    width: 28px;
    background: transparent;
}
QComboBox QAbstractItemView {
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    background: #ffffff;
    selection-background-color: #eff6ff;
    selection-color: #1e293b;
    padding: 4px;
}
/* ─── CheckBox, Radio ─── */
QCheckBox, QRadioButton {
    font-size: 13px;
    color: #334155;
    spacing: 10px;
}
QCheckBox::indicator, QRadioButton::indicator {
    width: 18px;
    height: 18px;
    border: 2px solid #cbd5e1;
    border-radius: 4px;
    background: white;
}
QCheckBox::indicator:checked, QRadioButton::indicator:checked {
    background-color: #2563eb;
    border-color: #2563eb;
}
QCheckBox#tableCheckBox {
    margin-left: 12px;
}
/* ─── SpinBox ─── */
QDoubleSpinBox, QSpinBox {
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    padding: 8px 12px;
    font-size: 13px;
    background: #ffffff;
    color: #1e293b;
}
QDoubleSpinBox:focus, QSpinBox:focus {
    border-color: #3b82f6;
}
/* ─── ScrollBar ─── */
QScrollBar:vertical {
    width: 8px;
    background: transparent;
    border-radius: 4px;
    margin: 2px 2px 2px 0;
}
QScrollBar::handle:vertical {
    background: #cbd5e1;
    border-radius: 4px;
    min-height: 40px;
}
QScrollBar::handle:vertical:hover {
    background: #94a3b8;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0;
}
QScrollBar:horizontal {
    height: 8px;
    background: transparent;
    border-radius: 4px;
}
QScrollBar::handle:horizontal {
    background: #cbd5e1;
    border-radius: 4px;
    min-width: 40px;
}
QScrollBar::handle:horizontal:hover {
    background: #94a3b8;
}
/* ─── Метки ─── */
QLabel#sectionTitle {
    font-size: 17px;
    font-weight: 600;
    color: #1e293b;
    letter-spacing: 0.2px;
}
QLabel#cardTitle {
    font-size: 14px;
    font-weight: 600;
    color: #1e293b;
}
QLabel#subsectionLabel {
    font-size: 13px;
    font-weight: 600;
    color: #475569;
}
QLabel#warningLabel {
    font-size: 14px;
    color: #dc2626;
}
QLabel#stepLabel {
    font-size: 13px;
    color: #64748b;
}
QLabel#hintLabel {
    font-size: 12px;
    color: #64748b;
}
QLabel#infoHint {
    font-size: 12px;
    color: #2563eb;
    padding: 4px 0;
}
QLabel#badge {
    background-color: #eff6ff;
    color: #2563eb;
    border-radius: 12px;
    padding: 4px 10px;
    font-size: 12px;
    font-weight: 600;
}
QLabel#badgeGreen {
    background-color: #f0fdf4;
    color: #16a34a;
    border-radius: 12px;
    padding: 4px 10px;
    font-size: 12px;
    font-weight: 600;
}
QLabel#badgeRed {
    background-color: #fef2f2;
    color: #dc2626;
    border-radius: 12px;
    padding: 4px 10px;
    font-size: 12px;
    font-weight: 600;
}
/* ─── Список ─── */
QListWidget {
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    background: #ffffff;
    font-size: 13px;
    color: #1e293b;
}
QListWidget::item {
    padding: 10px 14px;
    border-bottom: 1px solid #f8fafc;
}
QListWidget::item:selected {
    background-color: #eff6ff;
    color: #1e293b;
}
/* ─── Дерево ─── */
QTreeWidget {
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    background: #ffffff;
    font-size: 13px;
    color: #1e293b;
}
QTreeWidget::item {
    padding: 6px 8px;
}
QTreeWidget::item:selected {
    background-color: #eff6ff;
    color: #1e293b;
}
/* ─── Панель редактирования (превью) ─── */
QFrame#editPanel {
    background-color: #f8fafc;
    border-left: 2px solid #e2e8f0;
}
QLabel#panelCaption {
    font-size: 11px;
    color: #64748b;
}
QLabel#panelTitle {
    font-size: 13px;
    font-weight: bold;
    color: #1e293b;
}
QLabel#panelReadOnly {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 6px;
    padding: 8px;
    font-size: 12px;
    color: #1e293b;
}
QLabel#panelHighlight {
    font-size: 22px;
    font-weight: bold;
    color: #2563eb;
    padding: 4px 0;
}
QPushButton#btnPanelClose {
    background: transparent;
    border: none;
    color: #94a3b8;
    font-size: 14px;
    font-weight: bold;
    padding: 0;
    min-width: 24px;
    max-width: 24px;
    min-height: 24px;
    max-height: 24px;
}
QPushButton#btnPanelClose:hover {
    color: #dc2626;
}
QLabel#dropZoneIcon {
    font-size: 32px;
}
/* ─── Прогресс ─── */
QProgressBar {
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    background: #f1f5f9;
    text-align: center;
    font-size: 12px;
    color: #64748b;
    max-height: 22px;
}
QProgressBar::chunk {
    background: qlineargradient(x1:0 y1:0 x2:1 y2:0, stop:0 #3b82f6, stop:1 #2563eb);
    border-radius: 6px;
}
"""
