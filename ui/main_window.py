"""
main_window.py — Главное окно приложения.
Содержит: меню, стек страниц, навигацию.
"""
import os

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QStackedWidget, QMenuBar, QMenu, QLabel, QPushButton,
    QFrame, QSizePolicy
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QAction, QFont, QIcon, QKeySequence, QShortcut, QPixmap, QPainter


STYLESHEET = """
QMainWindow {
    background-color: #f5f5f5;
}
QWidget#centralWidget {
    background-color: #f5f5f5;
}
/* ─── Заголовок ─── */
QWidget#headerBar {
    background-color: #1e293b;
    min-height: 48px;
    max-height: 48px;
}
QLabel#appTitle {
    color: #f8fafc;
    font-size: 15px;
    font-weight: 600;
    padding-left: 16px;
}
QLabel#pageTitle {
    color: #94a3b8;
    font-size: 13px;
    padding-right: 16px;
}
/* ─── Меню ─── */
QMenuBar {
    background-color: #1e293b;
    color: #cbd5e1;
    font-size: 13px;
    spacing: 2px;
    padding: 0 8px;
}
QMenuBar::item {
    background: transparent;
    padding: 6px 12px;
    border-radius: 4px;
}
QMenuBar::item:selected {
    background-color: #334155;
    color: #f8fafc;
}
QMenu {
    background-color: #1e293b;
    color: #e2e8f0;
    border: 1px solid #334155;
    border-radius: 6px;
    padding: 4px 0;
    font-size: 13px;
}
QMenu::item {
    padding: 8px 20px;
}
QMenu::item:selected {
    background-color: #334155;
    color: #f8fafc;
}
QMenu::separator {
    height: 1px;
    background: #334155;
    margin: 4px 0;
}
/* ─── Кнопки ─── */
QPushButton {
    font-size: 13px;
    border-radius: 6px;
    padding: 8px 18px;
    font-weight: 500;
}
QPushButton#btnPrimary {
    background-color: #2563eb;
    color: white;
    border: none;
}
QPushButton#btnPrimary:hover {
    background-color: #1d4ed8;
}
QPushButton#btnPrimary:disabled {
    background-color: #93c5fd;
    color: #dbeafe;
}
QPushButton#btnSecondary {
    background-color: #e2e8f0;
    color: #1e293b;
    border: none;
}
QPushButton#btnSecondary:hover {
    background-color: #cbd5e1;
}
QPushButton#btnDanger {
    background-color: #fee2e2;
    color: #dc2626;
    border: none;
}
QPushButton#btnDanger:hover {
    background-color: #fecaca;
}
QPushButton#btnBack {
    background-color: transparent;
    color: #64748b;
    border: 1px solid #cbd5e1;
    padding: 6px 14px;
}
QPushButton#btnBack:hover {
    background-color: #f1f5f9;
    color: #1e293b;
}
QPushButton#btnIcon {
    background-color: transparent;
    color: #64748b;
    border: none;
    padding: 4px 8px;
    font-size: 16px;
}
QPushButton#btnIcon:hover {
    color: #1e293b;
}
/* ─── Карточки ─── */
QFrame#card {
    background-color: white;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
}
/* ─── Поля ввода ─── */
QLineEdit, QTextEdit {
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 6px 10px;
    font-size: 13px;
    background: white;
    color: #1e293b;
}
QLineEdit:focus, QTextEdit:focus {
    border-color: #2563eb;
    outline: none;
}
/* ─── Таблицы ─── */
QTableWidget {
    border: 1px solid #e2e8f0;
    border-radius: 6px;
    background: white;
    gridline-color: #f1f5f9;
    font-size: 13px;
    color: #1e293b;
}
QTableWidget::item {
    padding: 6px 8px;
}
QTableWidget::item:selected {
    background-color: #eff6ff;
    color: #1e293b;
}
QHeaderView::section {
    background-color: #f8fafc;
    color: #64748b;
    font-weight: 600;
    font-size: 12px;
    padding: 8px;
    border: none;
    border-bottom: 1px solid #e2e8f0;
    border-right: 1px solid #e2e8f0;
}
/* ─── ComboBox ─── */
QComboBox {
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 6px 10px;
    font-size: 13px;
    background: white;
    color: #1e293b;
    min-width: 160px;
}
QComboBox:focus {
    border-color: #2563eb;
}
QComboBox::drop-down {
    border: none;
    width: 24px;
}
QComboBox QAbstractItemView {
    border: 1px solid #e2e8f0;
    border-radius: 6px;
    background: white;
    selection-background-color: #eff6ff;
    selection-color: #1e293b;
}
/* ─── CheckBox ─── */
QCheckBox {
    font-size: 13px;
    color: #1e293b;
    spacing: 8px;
}
QCheckBox::indicator {
    width: 16px;
    height: 16px;
    border: 2px solid #cbd5e1;
    border-radius: 4px;
    background: white;
}
QCheckBox::indicator:checked {
    background-color: #2563eb;
    border-color: #2563eb;
}
/* ─── SpinBox ─── */
QDoubleSpinBox, QSpinBox {
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 6px 10px;
    font-size: 13px;
    background: white;
    color: #1e293b;
}
QDoubleSpinBox:focus, QSpinBox:focus {
    border-color: #2563eb;
}
/* ─── RadioButton ─── */
QRadioButton {
    font-size: 13px;
    color: #1e293b;
    spacing: 8px;
}
QRadioButton::indicator {
    width: 16px;
    height: 16px;
}
/* ─── ScrollBar ─── */
QScrollBar:vertical {
    width: 8px;
    background: #f1f5f9;
    border-radius: 4px;
}
QScrollBar::handle:vertical {
    background: #cbd5e1;
    border-radius: 4px;
    min-height: 20px;
}
QScrollBar::handle:vertical:hover {
    background: #94a3b8;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0;
}
/* ─── Разделители ─── */
QFrame#separator {
    background-color: #e2e8f0;
    max-height: 1px;
}
/* ─── Метки ─── */
QLabel#sectionTitle {
    font-size: 16px;
    font-weight: 600;
    color: #1e293b;
}
QLabel#stepLabel {
    font-size: 13px;
    color: #64748b;
}
QLabel#badge {
    background-color: #eff6ff;
    color: #2563eb;
    border-radius: 10px;
    padding: 2px 8px;
    font-size: 12px;
    font-weight: 600;
}
QLabel#badgeGreen {
    background-color: #f0fdf4;
    color: #16a34a;
    border-radius: 10px;
    padding: 2px 8px;
    font-size: 12px;
    font-weight: 600;
}
QLabel#badgeRed {
    background-color: #fef2f2;
    color: #dc2626;
    border-radius: 10px;
    padding: 2px 8px;
    font-size: 12px;
    font-weight: 600;
}
/* ─── Список файлов ─── */
QListWidget {
    border: 1px solid #e2e8f0;
    border-radius: 6px;
    background: white;
    font-size: 13px;
    color: #1e293b;
}
QListWidget::item {
    padding: 8px 12px;
    border-bottom: 1px solid #f1f5f9;
}
QListWidget::item:selected {
    background-color: #eff6ff;
    color: #1e293b;
}
/* ─── Прогресс ─── */
QProgressBar {
    border: 1px solid #e2e8f0;
    border-radius: 6px;
    background: #f1f5f9;
    text-align: center;
    font-size: 12px;
    color: #64748b;
    max-height: 20px;
}
QProgressBar::chunk {
    background-color: #2563eb;
    border-radius: 5px;
}
"""


class MainWindow(QMainWindow):
    """Главное окно приложения."""

    # Сигналы навигации
    navigate_to = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Маршруты, Сборка")
        self.setMinimumSize(1100, 700)
        self.resize(1200, 750)
        self.setStyleSheet(STYLESHEET)
        self._set_window_icon()

        # Состояние приложения (передаётся между страницами)
        self.app_state = {
            "fileType": "main",       # "main" | "increase"
            "filePaths": [],          # загруженные XLS файлы
            "saveDir": None,          # папка сохранения
            "routes": [],             # распарсенные маршруты
            "uniqueProducts": [],     # уникальные продукты из файлов
            "filteredRoutes": [],     # маршруты после фильтрации/исключения
            "routeCategory": "ШК",    # "ШК" | "СД" для округления
            "sortAsc": False,         # сортировка маршрутов
        }

        self._build_ui()
        self._build_menu()
        self._build_shortcuts()

    def _set_window_icon(self):
        """Устанавливает иконку окна (сборщик по маршруту): .ico для системы/сборки, иначе SVG."""
        base = os.path.dirname(os.path.abspath(__file__))
        res = os.path.join(base, "..", "resources")
        ico_path = os.path.join(res, "app.ico")
        svg_path = os.path.join(res, "app.svg")
        try:
            if os.path.isfile(ico_path):
                self.setWindowIcon(QIcon(ico_path))
                return
        except Exception:
            pass
        try:
            from PyQt6.QtSvg import QSvgRenderer
            if os.path.isfile(svg_path):
                renderer = QSvgRenderer(svg_path)
                icon = QIcon()
                for size in (16, 32, 48, 64):
                    pm = QPixmap(size, size)
                    pm.fill(Qt.GlobalColor.transparent)
                    painter = QPainter(pm)
                    renderer.render(painter)
                    painter.end()
                    icon.addPixmap(pm)
                self.setWindowIcon(icon)
        except Exception:
            pass

    # ─────────────────────────── UI ───────────────────────────────────

    def _build_ui(self):
        central = QWidget()
        central.setObjectName("centralWidget")
        self.setCentralWidget(central)

        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        # Заголовок
        self.header_bar = self._make_header()
        root_layout.addWidget(self.header_bar)

        # Стек страниц
        self.stack = QStackedWidget()
        root_layout.addWidget(self.stack)

    def _make_header(self) -> QWidget:
        bar = QWidget()
        bar.setObjectName("headerBar")
        bar.setFixedHeight(48)
        lay = QHBoxLayout(bar)
        lay.setContentsMargins(0, 0, 16, 0)
        lay.setSpacing(0)

        self.lbl_title = QLabel("Маршруты, Сборка")
        self.lbl_title.setObjectName("appTitle")
        lay.addWidget(self.lbl_title)

        lay.addStretch()

        self.lbl_page = QLabel("")
        self.lbl_page.setObjectName("pageTitle")
        lay.addWidget(self.lbl_page)

        return bar

    def _build_menu(self):
        """Строит меню-бар."""
        mb = self.menuBar()
        mb.setNativeMenuBar(False)

        # Меню «Файл»
        file_menu = mb.addMenu("Файл")
        act_new = QAction("Новая обработка", self)
        act_new.triggered.connect(self._on_new_session)
        file_menu.addAction(act_new)
        file_menu.addSeparator()
        act_process = QAction("Обработать файлы", self)
        act_process.triggered.connect(lambda: self.navigate_to.emit("home"))
        act_process.setShortcut(QKeySequence("Ctrl+O"))
        act_process.setToolTip("Загрузка и обработка XLS-файлов маршрутов")
        file_menu.addAction(act_process)
        act_labels = QAction("Этикетки", self)
        act_labels.triggered.connect(lambda: self.navigate_to.emit("labels"))
        act_labels.setToolTip("Создание этикеток XLS по шаблонам")
        file_menu.addAction(act_labels)
        file_menu.addSeparator()
        act_exit = QAction("Выход", self)
        act_exit.triggered.connect(self.close)
        file_menu.addAction(act_exit)

        # Меню «Справочники»
        ref_menu = mb.addMenu("Справочники")
        act_depts = QAction("Отделы и продукты\tCtrl+D", self)
        act_depts.setShortcut(QKeySequence("Ctrl+D"))
        act_depts.triggered.connect(lambda: self.navigate_to.emit("departments"))
        ref_menu.addAction(act_depts)

        act_templates = QAction("Шаблоны\tCtrl+T", self)
        act_templates.setShortcut(QKeySequence("Ctrl+T"))
        act_templates.triggered.connect(lambda: self.navigate_to.emit("templates"))
        ref_menu.addAction(act_templates)

        act_products = QAction("Продукты\tCtrl+P", self)
        act_products.setShortcut(QKeySequence("Ctrl+P"))
        act_products.setToolTip("Справочник продуктов (округление, шаблоны)")
        act_products.triggered.connect(lambda: self.navigate_to.emit("products"))
        ref_menu.addAction(act_products)

        # Меню «Настройки»
        settings_menu = mb.addMenu("Настройки")
        act_settings = QAction("Настройки Шт", self)
        act_settings.triggered.connect(lambda: self.navigate_to.emit("settings"))
        settings_menu.addAction(act_settings)

        # Меню «Помощь»
        help_menu = mb.addMenu("Помощь")
        act_shortcuts = QAction("Горячие клавиши", self)
        act_shortcuts.triggered.connect(self._show_shortcuts_help)
        help_menu.addAction(act_shortcuts)

    # ─────────────────────────── Горячие клавиши ────────────────────────

    def _build_shortcuts(self):
        """Регистрирует глобальные горячие клавиши.

        ВАЖНО: emit() принимает только существующие имена страниц.
        Ctrl+S, F5, Escape убраны — они отправляли несуществующие имена
        ('generate', 'refresh', 'back') и вызывали ошибки.
        """
        # Ctrl+O — обработка файлов
        sc_open = QShortcut(QKeySequence("Ctrl+O"), self)
        sc_open.activated.connect(lambda: self.navigate_to.emit("home"))
        sc_open.setWhatsThis("Обработка файлов (загрузка XLS)")
        # Ctrl+P — Продукты
        sc_prod = QShortcut(QKeySequence("Ctrl+P"), self)
        sc_prod.activated.connect(lambda: self.navigate_to.emit("products"))
        sc_prod.setWhatsThis("Справочник продуктов")

        # Ctrl+D — Отделы и продукты
        sc_dept = QShortcut(QKeySequence("Ctrl+D"), self)
        sc_dept.activated.connect(lambda: self.navigate_to.emit("departments"))
        sc_dept.setWhatsThis("Открыть Отделы и продукты")

        # Ctrl+T — Шаблоны
        sc_tmpl = QShortcut(QKeySequence("Ctrl+T"), self)
        sc_tmpl.activated.connect(lambda: self.navigate_to.emit("templates"))
        sc_tmpl.setWhatsThis("Открыть Шаблоны")

    def _show_shortcuts_help(self):
        from PyQt6.QtWidgets import QMessageBox
        QMessageBox.information(
            self, "Горячие клавиши",
            "Горячие клавиши приложения:\n\n"
            "  Ctrl+O    —  Обработка файлов (загрузка XLS)\n"
            "  Ctrl+D    —  Отделы и продукты\n"
            "  Ctrl+T    —  Шаблоны\n"
            "  Ctrl+P    —  Продукты\n"
        )

    # ─────────────────────────── Навигация ────────────────────────────────────

    def set_page_title(self, title: str):    self.lbl_page.setText(title)

    def _on_new_session(self):
        """Сбрасывает состояние и возвращает на главную страницу."""
        # Обновляем существующий словарь (страницы держат ссылку на него)
        self.app_state.update({
            "fileType": "main",
            "filePaths": [],
            "saveDir": None,
            "routes": [],
            "uniqueProducts": [],
            "filteredRoutes": [],
            "routeCategory": "ШК",
            "sortAsc": False,
        })
        self.navigate_to.emit("dashboard")
