"""
main_window.py — Главное окно приложения.
Содержит: меню, стек страниц, навигацию.
"""
import os
import sys

# Если файл запущен напрямую (не через app.py), добавить корень проекта в sys.path,
# чтобы работали импорты вида `from ui...` и `from core...`.
if __name__ == "__main__" and __package__ is None:
    ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if ROOT_DIR not in sys.path:
        sys.path.insert(0, ROOT_DIR)

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QStackedWidget, QMenuBar, QMenu, QLabel, QPushButton,
    QFrame, QSizePolicy, QTabBar,
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QAction, QFont, QKeySequence

from core import data_store


class MainWindow(QMainWindow):
    """Главное окно приложения."""

    # Сигналы навигации
    navigate_to = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Маршруты, Сборка")
        self.setMinimumSize(1000, 620)
        self.resize(1100, 700)

        # Состояние приложения (передаётся между страницами). Папка сохранения загружается из хранилища.
        save_dir = data_store.get_setting("defaultSaveDir")
        self.app_state = {
            "fileType": "main",       # "main" | "increase"
            "filePaths": [],          # загруженные XLS файлы
            "saveDir": save_dir,      # папка сохранения (из настроек для следующих запусков)
            "routes": [],             # распарсенные маршруты
            "uniqueProducts": [],     # уникальные продукты из файлов
            "filteredRoutes": [],     # маршруты после фильтрации/исключения
            "routeCategory": "ШК",    # "ШК" | "СД" для округления
            "sortAsc": True,          # сортировка маршрутов по возрастанию
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
        self.stack.setMinimumSize(900, 500)
        root_layout.addWidget(self.stack)

        # Строка состояния (как в Excel)
        self.statusBar().showMessage("Готово")

        # Колбэк для вывода сообщений в статус (страницы вызывают app_state["set_status"](msg))
        self.app_state["set_status"] = lambda msg, t=5000: self.statusBar().showMessage(msg, t)

    # Порядок вкладок ленты (как в Excel) — должен совпадать с порядком страниц в стеке
    RIBBON_PAGES = ["dashboard", "home", "labels", "preview_general", "preview_dept"]
    RIBBON_LABELS = ["Главная", "Обработка файлов", "Этикетки", "Общие маршруты", "По отделам"]

    def _make_header(self) -> QWidget:
        bar = QWidget()
        bar.setObjectName("headerBar")
        bar.setFixedHeight(28)
        lay = QHBoxLayout(bar)
        lay.setContentsMargins(6, 0, 6, 0)
        lay.setSpacing(0)

        # Вкладки ленты (как в Excel)
        self.ribbon_tabs = QTabBar()
        self.ribbon_tabs.setObjectName("ribbonTabBar")
        self.ribbon_tabs.setDocumentMode(True)
        self.ribbon_tabs.setExpanding(False)
        for label in self.RIBBON_LABELS:
            self.ribbon_tabs.addTab(label)
        self.ribbon_tabs.setCurrentIndex(0)
        self.ribbon_tabs.currentChanged.connect(self._on_ribbon_tab_changed)
        lay.addWidget(self.ribbon_tabs)

        self._update_routes_dependent_tabs()
        lay.addStretch()

        # Кнопка подсказки справа в ленте
        self.btn_page_hint = QPushButton("!")
        self.btn_page_hint.setObjectName("pageHintIcon")
        self.btn_page_hint.setToolTip("")
        self._page_hint_long = ""
        self.btn_page_hint.clicked.connect(self._show_page_hint_long)
        lay.addWidget(self.btn_page_hint)

        return bar

    def _on_ribbon_tab_changed(self, index: int):
        if 0 <= index < len(self.RIBBON_PAGES):
            if not self.ribbon_tabs.isTabEnabled(index):
                return
            self.navigate_to.emit(self.RIBBON_PAGES[index])

    def _update_routes_dependent_tabs(self) -> None:
        """Включает/выключает вкладки Этикетки, Общие маршруты, По отделам — только при наличии маршрутов."""
        routes = self.app_state.get("filteredRoutes") or self.app_state.get("routes") or []
        active = sum(1 for r in routes if not r.get("excluded"))
        has_routes = active > 0
        hint_disabled = "Сначала добавьте информацию о маршрутах"
        for i, name in enumerate(self.RIBBON_PAGES):
            if name in ("labels", "preview_general", "preview_dept"):
                self.ribbon_tabs.setTabEnabled(i, has_routes)
                self.ribbon_tabs.setTabToolTip(i, hint_disabled if not has_routes else "")

    def set_ribbon_page(self, page_name: str):
        """Устанавливает активную вкладку ленты по имени страницы (вызывается из app.navigate)."""
        if page_name in self.RIBBON_PAGES:
            idx = self.RIBBON_PAGES.index(page_name)
            if self.ribbon_tabs.currentIndex() != idx:
                self.ribbon_tabs.blockSignals(True)
                self.ribbon_tabs.setCurrentIndex(idx)
                self.ribbon_tabs.blockSignals(False)

    def _build_menu(self):
        """Строит многоуровневое меню-бар с вложенными подменю."""
        mb = self.menuBar()
        mb.setNativeMenuBar(False)

        # ── Файл (Справочники, Настройки, Помощь) ───────────────────────────
        file_menu = mb.addMenu("Файл")

        ref_sub = file_menu.addMenu("Справочники")
        data_sub = ref_sub.addMenu("Данные")
        act_depts = QAction("Отделы и продукты\tCtrl+D", self)
        act_depts.setShortcut(QKeySequence("Ctrl+D"))
        act_depts.triggered.connect(lambda: self.navigate_to.emit("departments"))
        data_sub.addAction(act_depts)
        act_products = QAction("Продукты\tCtrl+P", self)
        act_products.setShortcut(QKeySequence("Ctrl+P"))
        act_products.triggered.connect(lambda: self.navigate_to.emit("products"))
        data_sub.addAction(act_products)
        act_templates = QAction("Шаблоны Excel\tCtrl+T", self)
        act_templates.setShortcut(QKeySequence("Ctrl+T"))
        act_templates.triggered.connect(lambda: self.navigate_to.emit("templates"))
        ref_sub.addAction(act_templates)

        settings_sub = file_menu.addMenu("Настройки")
        act_file_params = QAction("Параметры создания файлов", self)
        act_file_params.triggered.connect(self._open_file_creation_settings)
        settings_sub.addAction(act_file_params)
        act_quantity = QAction("Настройки Количества", self)
        act_quantity.triggered.connect(self._open_quantity_settings)
        settings_sub.addAction(act_quantity)

        help_sub = file_menu.addMenu("Помощь")
        act_shortcuts = QAction("Горячие клавиши", self)
        act_shortcuts.triggered.connect(self._show_shortcuts_help)
        help_sub.addAction(act_shortcuts)
        act_about = QAction("О программе", self)
        act_about.triggered.connect(self._show_about)
        help_sub.addAction(act_about)

        file_menu.addSeparator()
        act_exit = QAction("Выход", self)
        act_exit.triggered.connect(self.close)
        file_menu.addAction(act_exit)

    # ─────────────────────────── Горячие клавиши ────────────────────────

    def _build_shortcuts(self):
        """Горячие клавиши заданы у QAction в меню (Ctrl+O, Ctrl+L, Ctrl+D и т.д.)."""
        pass

    def _open_file_creation_settings(self):
        try:
            from ui.pages.file_creation_settings_dialog import open_file_creation_settings_dialog
            open_file_creation_settings_dialog(self)
        except Exception:
            import traceback
            import logging
            logging.getLogger("app").exception("file_creation_settings")

    def _open_quantity_settings(self):
        try:
            from ui.pages.quantity_settings_dialog import open_quantity_settings_dialog
            open_quantity_settings_dialog(self, self.app_state)
        except Exception:
            import traceback
            import logging
            logging.getLogger("app").exception("quantity_settings")

    def _show_shortcuts_help(self):
        from PyQt6.QtWidgets import QMessageBox
        text = (
            "Горячие клавиши приложения:\n\n"
            "  Ctrl+D    —  Отделы и продукты\n"
            "  Ctrl+T    —  Шаблоны\n"
            "  Ctrl+P    —  Продукты\n"
            "  Ctrl+F    —  Фокус на поиск (в предпросмотре и настройках)\n\n"
            "  Escape    —  Закрыть модальное окно или панель редактирования"
        )
        mb = QMessageBox(self)
        mb.setWindowTitle("Горячие клавиши")
        mb.setText(text)
        mb.setTextFormat(Qt.TextFormat.PlainText)
        mb.setIcon(QMessageBox.Icon.Information)
        mb.setStandardButtons(QMessageBox.StandardButton.Ok)
        mb.exec()

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

    def set_page_title(self, title: str):
        """Заголовок страницы отображается вкладкой ленты; метод оставлен для совместимости."""
        pass

    def set_page_hint(self, hint_short: str, hint_long: str = ""):
        """Краткая подсказка — при наведении; подробная — при нажатии на «!»."""
        self.btn_page_hint.setToolTip(hint_short)
        self._page_hint_long = hint_long or hint_short
        self.btn_page_hint.setVisible(bool(hint_short or hint_long))

    def _show_page_hint_long(self):
        if not self._page_hint_long:
            return
        from PyQt6.QtWidgets import QMessageBox
        mb = QMessageBox(self)
        mb.setWindowTitle("Инструкция")
        mb.setText(self._page_hint_long)
        mb.setTextFormat(Qt.TextFormat.PlainText)
        mb.setIcon(QMessageBox.Icon.Information)
        mb.setStandardButtons(QMessageBox.StandardButton.Ok)
        mb.exec()

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
            "sortAsc": True,
        })
        self.navigate_to.emit("dashboard")


if __name__ == "__main__":
    from PyQt6.QtWidgets import QApplication

    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())
