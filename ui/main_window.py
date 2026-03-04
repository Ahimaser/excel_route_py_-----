"""
main_window.py — Главное окно приложения.
Содержит: меню, стек страниц, навигацию.
"""
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QStackedWidget, QMenuBar, QMenu, QLabel, QPushButton,
    QFrame, QSizePolicy
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QAction, QFont, QKeySequence, QShortcut

from ui.styles import STYLESHEET


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

        # Статусная строка
        self.statusBar().showMessage("Готов")

        # Колбэк для вывода сообщений в статус (страницы вызывают app_state["set_status"](msg))
        self.app_state["set_status"] = lambda msg, t=5000: self.statusBar().showMessage(msg, t)

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
        act_new.setToolTip("Сбросить данные и вернуться на главную страницу")
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
        act_exit.setToolTip("Закрыть приложение")
        act_exit.triggered.connect(self.close)
        file_menu.addAction(act_exit)

        # Меню «Справочники»
        ref_menu = mb.addMenu("Справочники")
        act_depts = QAction("Отделы и продукты\tCtrl+D", self)
        act_depts.setShortcut(QKeySequence("Ctrl+D"))
        act_depts.setToolTip("Управление отделами, подотделами и привязкой продуктов")
        act_depts.triggered.connect(lambda: self.navigate_to.emit("departments"))
        ref_menu.addAction(act_depts)

        act_templates = QAction("Шаблоны\tCtrl+T", self)
        act_templates.setShortcut(QKeySequence("Ctrl+T"))
        act_templates.setToolTip("Настройка столбцов Excel-файлов по отделам")
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
        act_settings.setToolTip("Отображение в штуках, округление по продуктам и ШК/СД")
        act_settings.triggered.connect(lambda: self.navigate_to.emit("settings"))
        settings_menu.addAction(act_settings)

        # Меню «Помощь»
        help_menu = mb.addMenu("Помощь")
        act_shortcuts = QAction("Горячие клавиши", self)
        act_shortcuts.setToolTip("Показать список горячих клавиш")
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
        # Ctrl+L — Этикетки
        sc_labels = QShortcut(QKeySequence("Ctrl+L"), self)
        sc_labels.activated.connect(lambda: self.navigate_to.emit("labels"))
        sc_labels.setWhatsThis("Открыть Этикетки")

    def _show_shortcuts_help(self):
        from PyQt6.QtWidgets import QMessageBox
        QMessageBox.information(
            self, "Горячие клавиши",
            "Горячие клавиши приложения:\n\n"
            "  Ctrl+O    —  Обработка файлов (загрузка XLS)\n"
            "  Ctrl+L    —  Этикетки\n"
            "  Ctrl+D    —  Отделы и продукты\n"
            "  Ctrl+T    —  Шаблоны\n"
            "  Ctrl+P    —  Продукты\n"
            "  Ctrl+F    —  Фокус на поиск (в предпросмотре и настройках)\n\n"
            "  Escape    —  Закрыть модальное окно или панель редактирования"
        )

    def _show_about(self):
        from PyQt6.QtWidgets import QMessageBox
        try:
            from version import VERSION
        except ImportError:
            VERSION = "?"
        from core.data_store import get_app_data_dir
        data_dir = str(get_app_data_dir())
        QMessageBox.about(
            self, "О программе",
            f"<h3>Маршруты, Сборка</h3>"
            f"<p>Версия {VERSION}</p>"
            f"<p>Обработка маршрутных XLS файлов, генерация отчётов и этикеток.</p>"
            f"<p><small>Данные: {data_dir}</small></p>"
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
