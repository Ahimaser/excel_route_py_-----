"""
styles.py — Стили интерфейса с акцентом в стиле stail blue theme.
Основа остаётся «офисной», но палитра и элементы ближе к современному web‑UI.
"""

# Палитра в духе stail blue_theme / Material
_BG_MAIN = "#F5F7FB"           # Общий фон (светлый, слегка голубоватый)
_BG_SURFACE = "#FFFFFF"        # Фон панелей, карточек
_BG_CARD = "#FFFFFF"           # Карточки
_RIBBON_BG = "#FFFFFF"         # Верхняя панель / лента
_BORDER = "#D6DFEA"            # Границы, разделители
_BORDER_LIGHT = "#E3E9F3"      # Светлые границы
_ACCENT = "#0085DB"            # Основной синий (var(--mat-sys-primary))
_ACCENT_HOVER = "#006FB5"      # Hover для синей кнопки
_ACCENT_LIGHT = "rgba(0, 133, 219, 0.10)"  # Фон выделения / бейджей
_ACCENT_RIpple = "rgba(0, 133, 219, 0.18)" # Ripple / нажатие
_EXCEL_SELECT = "#0085DB"      # Цвет выделения строк
_EXCEL_SELECT_BG = "rgba(0, 133, 219, 0.14)"
_TEXT = "#111827"              # Основной текст (почти чёрный)
_TEXT_MUTED = "#6B7280"        # Вторичный текст (серый)
_TEXT_LIGHT = "#9CA3AF"        # Приглушённый
_SUCCESS = "#4BD08B"           # Успех (как в stail)
_DANGER = "#DC2626"            # Ошибка
_DANGER_LIGHT = "rgba(220, 38, 38, 0.10)"
_WARNING = "#F8C076"
_TABLE_HEADER_BG = "#EEF2FF"   # Заголовок таблицы со слабым синим тоном
_STATUS_BAR_BG = "#E5E7EB"     # Строка состояния
_MENU_BAR_BG = "#FFFFFF"       # Светлая верхняя полоса

# Публичные константы для инлайн-стилей
ACCENT = _ACCENT
ACCENT_LIGHT = _ACCENT_LIGHT

STYLESHEET = f"""
/* ─── База: стиль окна Excel ─── */
QMainWindow, QDialog, QWidget {{
    background-color: {_BG_MAIN};
    font-family: "Segoe UI", "Calibri", sans-serif;
    font-size: 11pt;
}}
QWidget#centralWidget {{
    background-color: {_BG_MAIN};
}}

/* ─── Лента (вкладки как в Excel) ─── */
QWidget#headerBar {{
    background-color: {_RIBBON_BG};
    min-height: 32px;
    max-height: 32px;
    border: none;
    border-bottom: 1px solid {_BORDER};
}}
QTabBar::tab {{
    background-color: transparent;
    color: {_TEXT};
    padding: 6px 14px;
    margin-right: 2px;
    border: 1px solid transparent;
    border-bottom: none;
    font-size: 11pt;
}}
QTabBar::tab:selected {{
    background-color: {_BG_MAIN};
    border: 1px solid {_BORDER};
    border-bottom: 1px solid {_BG_MAIN};
    margin-bottom: -1px;
}}
QTabBar::tab:hover:!selected {{
    background-color: {_BORDER_LIGHT};
}}
QTabBar::tab:!selected {{
    color: {_TEXT_MUTED};
}}
QTabBar#ribbonTabBar {{
    background: transparent;
    border: none;
}}
/* Кнопка подсказки в ленте */
QLabel#pageHintIcon, QPushButton#pageHintIcon {{
    background-color: transparent;
    color: {_TEXT};
    font-size: 11pt;
    padding: 4px 10px;
    border: 1px solid {_BORDER};
    border-radius: 2px;
    min-width: 28px;
}}
QPushButton#pageHintIcon:focus {{ outline: none; }}
QLabel#pageHintIcon:hover, QPushButton#pageHintIcon:hover {{
    background-color: {_BORDER_LIGHT};
}}
QDialog QLabel#pageHintIcon, QDialog QPushButton#pageHintIcon {{
    background-color: {_ACCENT_LIGHT};
    color: {_ACCENT};
    border-color: {_ACCENT};
}}

/* ─── Меню (как в Excel — светлая полоса) ─── */
QMenuBar {{
    background-color: {_MENU_BAR_BG};
    color: {_TEXT};
    font-size: 11pt;
    padding: 2px 4px;
    border-bottom: 1px solid {_BORDER};
}}
QMenuBar::item {{
    background: transparent;
    padding: 6px 12px;
    color: {_TEXT};
}}
QMenuBar::item:selected {{
    background-color: {_EXCEL_SELECT};
    color: #FFFFFF;
}}
QMenuBar::item:disabled {{
    color: {_TEXT_LIGHT};
}}
QMenu {{
    background-color: {_BG_CARD};
    color: {_TEXT};
    border: 1px solid {_BORDER};
    border-radius: 12px;
    padding: 8px 0;
    font-size: 14px;
}}
QMenu::item {{
    padding: 12px 24px;
    color: {_TEXT};
}}
QMenu::item:selected {{
    background-color: {_ACCENT_LIGHT};
    color: {_ACCENT};
}}
QMenu::item:disabled {{
    color: {_TEXT_LIGHT};
}}
QMenu::separator {{
    height: 1px;
    background: {_BORDER_LIGHT};
    margin: 6px 16px;
}}

/* ─── Кнопки (плоские, как в Excel) ─── */
QPushButton {{
    font-size: 11pt;
    font-weight: 400;
    border-radius: 2px;
    padding: 6px 14px;
    min-height: 28px;
    min-width: 80px;
}}
QPushButton#btnPrimary {{
    background-color: {_ACCENT};
    color: #FFFFFF;
    border: none;
}}
QPushButton#btnPrimary:hover {{
    background-color: {_ACCENT_HOVER};
}}
QPushButton#btnPrimary:pressed {{
    background-color: {_ACCENT_HOVER};
}}
QPushButton#btnPrimary:disabled {{
    background-color: {_TEXT_LIGHT};
    color: rgba(255,255,255,0.7);
}}
QPushButton#btnSecondary {{
    background-color: {_BG_CARD};
    color: {_ACCENT};
    border: 1px solid {_BORDER};
}}
QPushButton#btnSecondary:hover {{
    background-color: {_ACCENT_LIGHT};
    border-color: {_ACCENT};
}}
QPushButton#btnSecondary:pressed {{
    background-color: {_ACCENT_RIpple};
}}
QPushButton#btnDanger {{
    background-color: {_DANGER_LIGHT};
    color: {_DANGER};
    border: 1px solid #F5C6C3;
}}
QPushButton#btnDanger:hover {{
    background-color: #FAD2CF;
}}
QPushButton#btnBack {{
    background-color: transparent;
    color: {_TEXT_MUTED};
    border: 1px solid {_BORDER};
    padding: 6px 14px;
}}
QPushButton#btnBack:hover {{
    background-color: {_BG_SURFACE};
    color: {_TEXT};
}}
QPushButton#btnIcon {{
    background-color: transparent;
    color: {_TEXT_MUTED};
    border: none;
    padding: 4px 10px;
    font-size: 16px;
    border-radius: 14px;
    min-height: 28px;
}}
QPushButton#btnIcon:hover {{
    background-color: rgba(0,0,0,0.06);
    color: {_TEXT};
}}
QPushButton#btnIconDanger {{
    background-color: transparent;
    color: {_DANGER};
    border: none;
    padding: 4px 10px;
    font-size: 16px;
    border-radius: 14px;
}}
QPushButton#btnIconDanger:hover {{
    background-color: {_DANGER_LIGHT};
}}

/* ─── Карточки (плоские, как панели Excel) ─── */
QFrame#card {{
    background-color: {_BG_CARD};
    border: 1px solid {_BORDER};
    border-radius: 2px;
    padding: 12px;
}}

/* ─── Вкладки внутри страниц (QTabWidget) ─── */
QTabWidget::pane {{
    background-color: {_BG_CARD};
    border: 1px solid {_BORDER};
    border-radius: 0;
    margin-top: -1px;
    padding: 12px;
}}
QTabWidget::tab-bar {{
    alignment: left;
}}
QTabWidget QTabBar::tab {{
    background: {_RIBBON_BG};
    color: {_TEXT};
    padding: 6px 16px;
    margin-right: 2px;
    border: 1px solid {_BORDER};
    border-bottom: none;
}}
QTabWidget QTabBar::tab:selected {{
    background: {_BG_CARD};
    border-bottom: 1px solid {_BG_CARD};
    margin-bottom: -1px;
}}

/* ─── GroupBox ─── */
QGroupBox {{
    font-size: 14px;
    font-weight: 500;
    color: {_TEXT};
    border: 1px solid {_BORDER};
    border-radius: 12px;
    margin-top: 16px;
    padding: 20px 20px 12px 20px;
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 16px;
    padding: 0 8px;
    background-color: {_BG_CARD};
    color: {_TEXT_MUTED};
}}

/* ─── Поля ввода (как в Excel) ─── */
QLineEdit, QTextEdit {{
    border: 1px solid {_BORDER};
    border-radius: 0;
    padding: 4px 8px;
    font-size: 11pt;
    background: {_BG_MAIN};
    color: {_TEXT};
    selection-background-color: {_EXCEL_SELECT};
    selection-color: #FFFFFF;
    min-width: 8em;
}}
QLineEdit:focus, QTextEdit:focus {{
    border: 1px solid {_ACCENT};
}}
QLineEdit:hover, QTextEdit:hover {{
    border-color: {_TEXT_LIGHT};
}}

/* ─── Таблицы (сетка как в Excel) ─── */
QTableWidget, QTableView {{
    border: 1px solid {_BORDER};
    border-radius: 0;
    background: {_BG_MAIN};
    gridline-color: {_BORDER};
    font-size: 11pt;
    color: {_TEXT};
}}
QTableWidget {{
    alternate-background-color: #FAFAFA;
}}
QTableView {{
    alternate-background-color: #FAFAFA;
}}
QTableWidget::item, QTableView::item {{
    padding: 10px 14px;
    min-height: 28px;
}}
QTableWidget::item:selected, QTableView::item:selected {{
    background-color: {_EXCEL_SELECT_BG};
    color: {_TEXT};
}}
QTableWidget::item:hover, QTableView::item:hover {{
    background-color: #E9EDF4;
}}
QHeaderView::section {{
    background-color: {_TABLE_HEADER_BG};
    color: {_TEXT};
    font-weight: 600;
    font-size: 11pt;
    padding: 10px 14px;
    border: 1px solid {_BORDER};
}}

/* ─── ComboBox (Material dropdown) ─── */
QComboBox {{
    border: 1px solid {_BORDER};
    border-radius: 0;
    padding: 4px 8px;
    font-size: 11pt;
    background: {_BG_MAIN};
    color: {_TEXT};
    min-width: 200px;
    min-height: 24px;
}}
QComboBox:focus {{
    border: 1px solid {_ACCENT};
}}
QComboBox:hover {{
    border-color: {_TEXT_LIGHT};
}}
QComboBox::drop-down {{
    border: none;
    width: 36px;
    background: transparent;
}}
QComboBox QAbstractItemView {{
    min-width: 220px;
    border: 1px solid {_BORDER};
    border-radius: 8px;
    background: {_BG_CARD};
    selection-background-color: {_ACCENT_LIGHT};
    selection-color: {_TEXT};
    padding: 8px;
}}
QComboBox QAbstractItemView::item {{
    min-height: 28px;
}}

/* ─── CheckBox, Radio (Material) ─── */
QCheckBox, QRadioButton {{
    font-size: 14px;
    color: {_TEXT};
    spacing: 12px;
}}
QCheckBox::indicator {{
    width: 22px;
    height: 22px;
    border: 2px solid {_BORDER_LIGHT};
    border-radius: 6px;
    background: {_BG_SURFACE};
}}
QCheckBox::indicator:hover {{
    border-color: {_ACCENT};
    background: {_BG_MAIN};
}}
QCheckBox::indicator:checked {{
    background-color: {_ACCENT};
    border-color: {_ACCENT};
    image: url("data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 14 14'><polyline points='2,7 5.5,10.5 12,3' stroke='white' stroke-width='2.2' fill='none' stroke-linecap='round' stroke-linejoin='round'/></svg>");
}}
QRadioButton::indicator {{
    width: 26px;
    height: 26px;
    border: 1px solid {_BORDER_LIGHT};
    border-radius: 13px;
    background: {_BG_SURFACE};
}}
QRadioButton::indicator:checked {{
    background-color: {_BG_SURFACE};
    border: 7px solid {_ACCENT};
}}

/* ─── SpinBox ─── */
QDoubleSpinBox, QSpinBox {{
    border: 1px solid {_BORDER};
    border-radius: 8px;
    padding: 12px 16px;
    font-size: 14px;
    background: {_BG_CARD};
    color: {_TEXT};
}}
QDoubleSpinBox:focus, QSpinBox:focus {{
    border: 2px solid {_ACCENT};
}}

/* ─── ScrollBar (Material minimal) ─── */
QScrollBar:vertical {{
    width: 8px;
    background: transparent;
    border-radius: 4px;
    margin: 4px 2px 4px 0;
}}
QScrollBar::handle:vertical {{
    background: {_TEXT_LIGHT};
    border-radius: 4px;
    min-height: 40px;
}}
QScrollBar::handle:vertical:hover {{
    background: {_TEXT_MUTED};
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0;
}}
QScrollBar:horizontal {{
    height: 8px;
    background: transparent;
    border-radius: 4px;
}}
QScrollBar::handle:horizontal {{
    background: {_TEXT_LIGHT};
    border-radius: 4px;
    min-width: 40px;
}}
QScrollBar::handle:horizontal:hover {{
    background: {_TEXT_MUTED};
}}

/* ─── Labels (Material typography) ─── */
QLabel#sectionTitle {{
    font-size: 22px;
    font-weight: 500;
    color: {_TEXT};
    letter-spacing: 0.15px;
}}
QLabel#cardTitle {{
    font-size: 16px;
    font-weight: 500;
    color: {_TEXT};
}}
QLabel#subsectionLabel {{
    font-size: 14px;
    font-weight: 500;
    color: {_TEXT_MUTED};
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
    padding: 8px 0;
}}
QLabel#badge {{
    background-color: {_ACCENT_LIGHT};
    color: {_ACCENT};
    border-radius: 16px;
    padding: 6px 14px;
    font-size: 12px;
    font-weight: 600;
}}
QLabel#badgeGreen {{
    background-color: #E8F5E9;
    color: {_SUCCESS};
    border-radius: 16px;
    padding: 6px 14px;
    font-size: 12px;
    font-weight: 600;
}}
QLabel#badgeRed {{
    background-color: {_DANGER_LIGHT};
    color: {_DANGER};
    border-radius: 16px;
    padding: 6px 14px;
    font-size: 12px;
    font-weight: 600;
}}

/* ─── List (как таблица Excel: сетка, чередование строк) ─── */
QListWidget {{
    border: 1px solid {_BORDER};
    border-radius: 0;
    background: {_BG_MAIN};
    font-size: 11pt;
    color: {_TEXT};
}}
QListWidget::item {{
    padding: 10px 14px;
    min-height: 28px;
}}
QListWidget::item:alternate {{
    background-color: #FAFAFA;
}}
QListWidget::item:selected {{
    background-color: {_EXCEL_SELECT_BG};
    color: {_TEXT};
}}
QListWidget::item:hover {{
    background-color: #E9EDF4;
}}

/* ─── Tree (как таблица Excel) ─── */
QTreeWidget {{
    border: 1px solid {_BORDER};
    border-radius: 0;
    background: {_BG_MAIN};
    font-size: 11pt;
    color: {_TEXT};
}}
QTreeWidget::item {{
    padding: 10px 14px;
    min-height: 28px;
}}
QTreeWidget::item:selected {{
    background-color: {_EXCEL_SELECT_BG};
    color: {_TEXT};
}}
QTreeWidget::item:hover {{
    background-color: #E9EDF4;
}}

/* ─── Edit panel (side sheet) ─── */
QFrame#editPanel {{
    background-color: {_BG_SURFACE};
    border-left: 1px solid {_BORDER};
}}
QLabel#panelCaption {{
    font-size: 12px;
    color: {_TEXT_MUTED};
}}
QLabel#panelTitle {{
    font-size: 14px;
    font-weight: 500;
    color: {_TEXT};
}}
QLabel#panelReadOnly {{
    background: {_BG_CARD};
    border: 1px solid {_BORDER};
    border-radius: 8px;
    padding: 12px;
    font-size: 14px;
    color: {_TEXT};
}}
QLabel#panelHighlight {{
    font-size: 24px;
    font-weight: 500;
    color: {_ACCENT};
    padding: 8px 0;
}}
QPushButton#btnPanelClose {{
    background: transparent;
    border: none;
    color: {_TEXT_LIGHT};
    font-size: 18px;
    font-weight: bold;
    padding: 0;
    min-width: 32px;
    max-width: 32px;
    min-height: 32px;
    max-height: 32px;
    border-radius: 16px;
}}
QPushButton#btnPanelClose:hover {{
    color: {_DANGER};
    background-color: {_DANGER_LIGHT};
}}
QLabel#dropZoneIcon {{
    font-size: 40px;
}}

/* ─── Progress (Material linear) ─── */
QProgressBar {{
    border: none;
    border-radius: 8px;
    background: {_BORDER_LIGHT};
    text-align: center;
    font-size: 12px;
    color: {_TEXT_MUTED};
    max-height: 8px;
}}
QProgressBar::chunk {{
    background-color: {_ACCENT};
    border-radius: 8px;
}}

/* ─── Separator ─── */
QFrame#separator {{
    background-color: {_BORDER};
    max-height: 1px;
}}

/* ─── Строка состояния (как в Excel) ─── */
QStatusBar {{
    font-size: 11pt;
    color: {_TEXT};
    padding: 4px 12px;
    background: {_STATUS_BAR_BG};
    border-top: 1px solid {_BORDER};
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
