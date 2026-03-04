"""
styles.py — Современные стили приложения.
Оптимизированы для работы с текстом и таблицами.
"""
# Цветовая палитра
_BG_MAIN = "#f8fafc"
_BG_CARD = "#ffffff"
_BG_HEADER = "#0f172a"
_BG_HEADER_ALT = "#1e293b"
_BORDER = "#e2e8f0"
_BORDER_LIGHT = "#f1f5f9"
_ACCENT = "#2563eb"
_ACCENT_HOVER = "#1d4ed8"
_TEXT = "#1e293b"
_TEXT_MUTED = "#64748b"
_TEXT_LIGHT = "#94a3b8"
_SUCCESS = "#16a34a"
_DANGER = "#dc2626"
_WARNING = "#d97706"

STYLESHEET = f"""
/* ─── База: современная светлая тема ─── */
QMainWindow, QDialog, QWidget {{
    background-color: {_BG_MAIN};
    font-family: "Segoe UI", "SF Pro Display", -apple-system, sans-serif;
}}
QWidget#centralWidget {{
    background-color: {_BG_MAIN};
}}

/* ─── Заголовок приложения ─── */
QWidget#headerBar {{
    background: qlineargradient(x1:0 y1:0 x2:0 y2:1, stop:0 {_BG_HEADER}, stop:1 {_BG_HEADER_ALT});
    min-height: 52px;
    max-height: 52px;
    border-bottom: 1px solid rgba(148, 163, 184, 0.15);
}}
QLabel#appTitle {{
    color: #f8fafc;
    font-size: 18px;
    font-weight: 600;
    letter-spacing: 0.5px;
    padding-left: 28px;
}}
QLabel#pageTitle {{
    color: #94a3b8;
    font-size: 13px;
    font-weight: 500;
    padding-right: 24px;
}}

/* ─── Меню ─── */
QMenuBar {{
    background: transparent;
    color: #cbd5e1;
    font-size: 13px;
    padding: 0 16px;
}}
QMenuBar::item {{
    background: transparent;
    padding: 10px 16px;
    border-radius: 6px;
}}
QMenuBar::item:selected {{
    background-color: rgba(255,255,255,0.1);
    color: #f8fafc;
}}
QMenu {{
    background-color: #1e293b;
    color: #e2e8f0;
    border: 1px solid #334155;
    border-radius: 10px;
    padding: 8px 0;
    font-size: 13px;
}}
QMenu::item {{
    padding: 12px 28px;
}}
QMenu::item:selected {{
    background-color: #334155;
    color: #f8fafc;
}}
QMenu::separator {{
    height: 1px;
    background: #334155;
    margin: 6px 16px;
}}

/* ─── Кнопки ─── */
QPushButton {{
    font-size: 13px;
    font-weight: 500;
    border-radius: 10px;
    padding: 10px 22px;
}}
QPushButton#btnPrimary {{
    background: qlineargradient(x1:0 y1:0 x2:0 y2:1, stop:0 #3b82f6, stop:1 {_ACCENT});
    color: white;
    border: none;
}}
QPushButton#btnPrimary:hover {{
    background: qlineargradient(x1:0 y1:0 x2:0 y2:1, stop:0 {_ACCENT}, stop:1 {_ACCENT_HOVER});
}}
QPushButton#btnPrimary:pressed {{
    background: {_ACCENT_HOVER};
}}
QPushButton#btnPrimary:disabled {{
    background: #94a3b8;
    color: #e2e8f0;
}}
QPushButton#btnSecondary {{
    background-color: {_BG_CARD};
    color: {_TEXT};
    border: 1px solid {_BORDER};
}}
QPushButton#btnSecondary:hover {{
    background-color: #f1f5f9;
    border-color: #cbd5e1;
}}
QPushButton#btnDanger {{
    background-color: #fef2f2;
    color: {_DANGER};
    border: 1px solid #fecaca;
}}
QPushButton#btnDanger:hover {{
    background-color: #fee2e2;
    border-color: #fca5a5;
}}
QPushButton#btnBack {{
    background-color: transparent;
    color: {_TEXT_MUTED};
    border: 1px solid {_BORDER};
    padding: 8px 18px;
}}
QPushButton#btnBack:hover {{
    background-color: #f8fafc;
    color: {_TEXT};
}}
QPushButton#btnIcon {{
    background-color: transparent;
    color: {_TEXT_MUTED};
    border: none;
    padding: 8px 12px;
    font-size: 16px;
    border-radius: 8px;
}}
QPushButton#btnIcon:hover {{
    background-color: rgba(0,0,0,0.05);
    color: {_TEXT};
}}
QPushButton#btnIconDanger {{
    background-color: transparent;
    color: {_DANGER};
    border: none;
    padding: 8px 12px;
    font-size: 16px;
    border-radius: 8px;
}}
QPushButton#btnIconDanger:hover {{
    background-color: #fef2f2;
    color: #b91c1c;
}}

/* ─── Карточки ─── */
QFrame#card {{
    background-color: {_BG_CARD};
    border: 1px solid {_BORDER};
    border-radius: 14px;
    padding: 6px;
}}

/* ─── Вкладки ─── */
QTabWidget::pane {{
    background-color: {_BG_CARD};
    border: 1px solid {_BORDER};
    border-radius: 12px;
    margin-top: -1px;
    padding: 16px;
}}
QTabBar::tab {{
    background-color: #f8fafc;
    color: {_TEXT_MUTED};
    padding: 12px 24px;
    margin-right: 4px;
    border: 1px solid {_BORDER};
    border-bottom: none;
    border-top-left-radius: 10px;
    border-top-right-radius: 10px;
    font-weight: 500;
}}
QTabBar::tab:selected {{
    background-color: {_BG_CARD};
    color: {_TEXT};
    border-color: {_BORDER};
    border-bottom: 1px solid {_BG_CARD};
}}
QTabBar::tab:hover:!selected {{
    background-color: #f1f5f9;
    color: #475569;
}}

/* ─── Группы ─── */
QGroupBox {{
    font-size: 13px;
    font-weight: 600;
    color: {_TEXT};
    border: 1px solid {_BORDER};
    border-radius: 12px;
    margin-top: 14px;
    padding: 18px 18px 10px 18px;
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 16px;
    padding: 0 10px;
    background-color: {_BG_CARD};
    color: #475569;
}}

/* ─── Поля ввода (текст) ─── */
QLineEdit, QTextEdit {{
    border: 1px solid {_BORDER};
    border-radius: 10px;
    padding: 10px 14px;
    font-size: 14px;
    background: {_BG_CARD};
    color: {_TEXT};
    selection-background-color: #dbeafe;
}}
QLineEdit:focus, QTextEdit:focus {{
    border-color: {_ACCENT};
}}
QLineEdit:hover, QTextEdit:hover {{
    border-color: #cbd5e1;
}}

/* ─── Таблицы (оптимизированы для данных) ─── */
QTableWidget, QTableView {{
    border: 1px solid {_BORDER};
    border-radius: 12px;
    background: {_BG_CARD};
    gridline-color: {_BORDER_LIGHT};
    font-size: 13px;
    color: {_TEXT};
}}
QTableWidget {{
    alternate-background-color: #fafbfc;
}}
QTableView {{
    alternate-background-color: #fafbfc;
}}
QTableWidget::item, QTableView {{
    padding: 10px 14px;
    min-height: 20px;
}}
QTableWidget::item:selected, QTableView::item:selected {{
    background-color: #eff6ff;
    color: {_TEXT};
}}
QTableWidget::item:hover, QTableView::item:hover {{
    background-color: #f8fafc;
}}
QHeaderView::section {{
    background-color: #f8fafc;
    color: #475569;
    font-weight: 600;
    font-size: 12px;
    padding: 12px 14px;
    border: none;
    border-bottom: 2px solid {_BORDER};
    border-right: 1px solid {_BORDER_LIGHT};
}}
/* ─── ComboBox ─── */
QComboBox {{
    border: 1px solid {_BORDER};
    border-radius: 10px;
    padding: 10px 14px;
    font-size: 13px;
    background: {_BG_CARD};
    color: {_TEXT};
    min-width: 160px;
}}
QComboBox:focus {{
    border-color: {_ACCENT};
}}
QComboBox:hover {{
    border-color: #cbd5e1;
}}
QComboBox::drop-down {{
    border: none;
    width: 32px;
    background: transparent;
}}
QComboBox QAbstractItemView {{
    border: 1px solid {_BORDER};
    border-radius: 10px;
    background: {_BG_CARD};
    selection-background-color: #eff6ff;
    selection-color: {_TEXT};
    padding: 8px;
}}

/* ─── CheckBox, Radio ─── */
QCheckBox, QRadioButton {{
    font-size: 13px;
    color: {_TEXT};
    spacing: 12px;
}}
QCheckBox::indicator, QRadioButton::indicator {{
    width: 20px;
    height: 20px;
    border: 2px solid #cbd5e1;
    border-radius: 5px;
    background: white;
}}
QCheckBox::indicator:checked, QRadioButton::indicator:checked {{
    background-color: {_ACCENT};
    border-color: {_ACCENT};
}}

/* ─── SpinBox ─── */
QDoubleSpinBox, QSpinBox {{
    border: 1px solid {_BORDER};
    border-radius: 10px;
    padding: 10px 14px;
    font-size: 13px;
    background: {_BG_CARD};
    color: {_TEXT};
}}
QDoubleSpinBox:focus, QSpinBox:focus {{
    border-color: {_ACCENT};
}}

/* ─── ScrollBar (минималистичный) ─── */
QScrollBar:vertical {{
    width: 10px;
    background: transparent;
    border-radius: 5px;
    margin: 4px 2px 4px 0;
}}
QScrollBar::handle:vertical {{
    background: #cbd5e1;
    border-radius: 5px;
    min-height: 50px;
}}
QScrollBar::handle:vertical:hover {{
    background: #94a3b8;
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0;
}}
QScrollBar:horizontal {{
    height: 10px;
    background: transparent;
    border-radius: 5px;
}}
QScrollBar::handle:horizontal {{
    background: #cbd5e1;
    border-radius: 5px;
    min-width: 50px;
}}
QScrollBar::handle:horizontal:hover {{
    background: #94a3b8;
}}

/* ─── Метки ─── */
QLabel#sectionTitle {{
    font-size: 20px;
    font-weight: 600;
    color: {_TEXT};
    letter-spacing: 0.3px;
}}
QLabel#cardTitle {{
    font-size: 15px;
    font-weight: 600;
    color: {_TEXT};
}}
QLabel#subsectionLabel {{
    font-size: 13px;
    font-weight: 600;
    color: #475569;
}}
QLabel#warningLabel {{
    font-size: 14px;
    color: {_DANGER};
}}
QLabel#stepLabel {{
    font-size: 14px;
    color: {_TEXT_MUTED};
    line-height: 1.5;
}}
QLabel#hintLabel {{
    font-size: 13px;
    color: {_TEXT_MUTED};
}}
QLabel#infoHint {{
    font-size: 13px;
    color: {_ACCENT};
    padding: 6px 0;
}}
QLabel#badge {{
    background-color: #eff6ff;
    color: {_ACCENT};
    border-radius: 14px;
    padding: 5px 12px;
    font-size: 12px;
    font-weight: 600;
}}
QLabel#badgeGreen {{
    background-color: #f0fdf4;
    color: {_SUCCESS};
    border-radius: 14px;
    padding: 5px 12px;
    font-size: 12px;
    font-weight: 600;
}}
QLabel#badgeRed {{
    background-color: #fef2f2;
    color: {_DANGER};
    border-radius: 14px;
    padding: 5px 12px;
    font-size: 12px;
    font-weight: 600;
}}

/* ─── Список ─── */
QListWidget {{
    border: 1px solid {_BORDER};
    border-radius: 12px;
    background: {_BG_CARD};
    font-size: 13px;
    color: {_TEXT};
}}
QListWidget::item {{
    padding: 12px 16px;
    border-bottom: 1px solid {_BORDER_LIGHT};
}}
QListWidget::item:selected {{
    background-color: #eff6ff;
    color: {_TEXT};
}}
QListWidget::item:hover {{
    background-color: #f8fafc;
}}

/* ─── Дерево ─── */
QTreeWidget {{
    border: 1px solid {_BORDER};
    border-radius: 12px;
    background: {_BG_CARD};
    font-size: 13px;
    color: {_TEXT};
}}
QTreeWidget::item {{
    padding: 8px 12px;
}}
QTreeWidget::item:selected {{
    background-color: #eff6ff;
    color: {_TEXT};
}}
QTreeWidget::item:hover {{
    background-color: #f8fafc;
}}

/* ─── Панель редактирования ─── */
QFrame#editPanel {{
    background-color: #f8fafc;
    border-left: 3px solid {_BORDER};
}}
QLabel#panelCaption {{
    font-size: 11px;
    color: {_TEXT_MUTED};
}}
QLabel#panelTitle {{
    font-size: 13px;
    font-weight: 600;
    color: {_TEXT};
}}
QLabel#panelReadOnly {{
    background: {_BG_CARD};
    border: 1px solid {_BORDER};
    border-radius: 8px;
    padding: 10px;
    font-size: 13px;
    color: {_TEXT};
}}
QLabel#panelHighlight {{
    font-size: 22px;
    font-weight: 600;
    color: {_ACCENT};
    padding: 6px 0;
}}
QPushButton#btnPanelClose {{
    background: transparent;
    border: none;
    color: {_TEXT_LIGHT};
    font-size: 16px;
    font-weight: bold;
    padding: 0;
    min-width: 28px;
    max-width: 28px;
    min-height: 28px;
    max-height: 28px;
}}
QPushButton#btnPanelClose:hover {{
    color: {_DANGER};
}}
QLabel#dropZoneIcon {{
    font-size: 36px;
}}

/* ─── Прогресс ─── */
QProgressBar {{
    border: 1px solid {_BORDER};
    border-radius: 10px;
    background: #f1f5f9;
    text-align: center;
    font-size: 12px;
    color: {_TEXT_MUTED};
    max-height: 24px;
}}
QProgressBar::chunk {{
    background: qlineargradient(x1:0 y1:0 x2:1 y2:0, stop:0 #3b82f6, stop:1 {_ACCENT});
    border-radius: 8px;
}}

/* ─── Разделитель ─── */
QFrame#separator {{
    background-color: {_BORDER};
    max-height: 1px;
}}

/* ─── ScrollArea ─── */
QScrollArea {{
    border: none;
    background: transparent;
}}
QScrollArea > QWidget > QWidget {{
    background: transparent;
}}
"""

