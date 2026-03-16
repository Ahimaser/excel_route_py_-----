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
from PyQt6.QtWidgets import QMessageBox

from core import data_store
from core.constants import FILE_TYPE_MAIN, ROUTE_CATEGORY_SHK


class MainWindow(QMainWindow):
    """Главное окно приложения."""

    # Сигналы навигации
    navigate_to = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Маршруты, Сборка")
        self.setMinimumSize(1100, 680)
        self.resize(1200, 750)

        save_dir = data_store.get_setting("defaultSaveDir")
        self.app_state = self._create_app_state(save_dir)

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
        self.stack.setMinimumSize(950, 550)
        root_layout.addWidget(self.stack, 1)

        # Строка состояния (как в Excel)
        self.statusBar().showMessage("Готово")

        # Колбэк для вывода сообщений в статус (страницы вызывают app_state["set_status"](msg))
        self.app_state["set_status"] = lambda msg, t=5000: self.statusBar().showMessage(msg, t)
        # Колбэк для обновления вкладок (страницы вызывают после создания файлов)
        self.app_state["_update_tabs"] = lambda: self._update_routes_dependent_tabs()

    # Порядок вкладок ленты (как в Excel) — должен совпадать с порядком страниц в стеке
    RIBBON_PAGES = ["dashboard", "home", "labels", "preview_general", "preview_dept"]
    RIBBON_LABELS = ["Главная", "Обработка файлов", "Этикетки", "Общие маршруты", "По отделам"]
    RIBBON_TAB_HINTS = {
        "dashboard": "Стартовая страница: описание, место сохранения, отчёт по последним маршрутам.",
        "home": "Загрузка XLS-файлов маршрутов (ШК/СД), выбор папки сохранения и даты.",
        "labels": "Создание файлов этикеток по отделам и продуктам.",
        "preview_general": "Таблица общих маршрутов: поиск, фильтр, редактирование, исключение. Переход к маршрутам по отделам.",
        "preview_dept": "Маршруты по отделам: вкладки отделов, создание XLS-файлов.",
    }

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
        hint_disabled = "Сначала обработайте файлы или откройте маршруты из истории"
        for i, name in enumerate(self.RIBBON_PAGES):
            if name == "labels":
                self.ribbon_tabs.setTabEnabled(i, has_routes)
                self.ribbon_tabs.setTabToolTip(i, hint_disabled if not has_routes else self.RIBBON_TAB_HINTS.get(name, ""))
            elif name in ("preview_general", "preview_dept"):
                self.ribbon_tabs.setTabEnabled(i, has_routes)
                self.ribbon_tabs.setTabToolTip(i, hint_disabled if not has_routes else self.RIBBON_TAB_HINTS.get(name, ""))
            else:
                self.ribbon_tabs.setTabToolTip(i, self.RIBBON_TAB_HINTS.get(name, ""))

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
        act_process = QAction("Обработать файлы\tCtrl+O", self)
        act_process.setShortcut(QKeySequence("Ctrl+O"))
        act_process.triggered.connect(lambda: self.navigate_to.emit("home"))
        file_menu.addAction(act_process)
        file_menu.addSeparator()

        ref_sub = file_menu.addMenu("Справочники")
        act_depts = QAction("Отделы и продукты\tCtrl+D", self)
        act_depts.setShortcut(QKeySequence("Ctrl+D"))
        act_depts.triggered.connect(lambda: self.navigate_to.emit("departments"))
        ref_sub.addAction(act_depts)
        act_products = QAction("Продукты\tCtrl+P", self)
        act_products.setShortcut(QKeySequence("Ctrl+P"))
        act_products.triggered.connect(lambda: self.navigate_to.emit("products"))
        ref_sub.addAction(act_products)

        settings_sub = file_menu.addMenu("Настройки")
        # Данные
        act_restore = QAction("Восстановить данные из резервной копии", self)
        act_restore.triggered.connect(self._on_restore_data)
        settings_sub.addAction(act_restore)
        act_templates = QAction("Шаблоны Excel\tCtrl+T", self)
        act_templates.setShortcut(QKeySequence("Ctrl+T"))
        act_templates.triggered.connect(lambda: self.navigate_to.emit("templates"))
        settings_sub.addAction(act_templates)
        settings_sub.addSeparator()
        # Создание файлов
        act_file_params = QAction("Параметры создания файлов", self)
        act_file_params.triggered.connect(self._open_file_creation_settings)
        settings_sub.addAction(act_file_params)
        act_save_mode = QAction("Режимы сохранения (ШК/СД)", self)
        act_save_mode.triggered.connect(self._open_save_mode_settings)
        settings_sub.addAction(act_save_mode)
        settings_sub.addSeparator()
        # Расчёты
        act_quantity = QAction("Настройки Количества", self)
        act_quantity.triggered.connect(self._open_quantity_settings)
        settings_sub.addAction(act_quantity)
        act_savings = QAction("Экономия", self)
        act_savings.triggered.connect(self._open_savings_settings)
        settings_sub.addAction(act_savings)
        settings_sub.addSeparator()
        # Внешний вид
        act_appearance = QAction("Оформление", self)
        act_appearance.triggered.connect(self._open_appearance_settings)
        settings_sub.addAction(act_appearance)

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

    def _create_app_state(self, save_dir: str | None) -> dict:
        """Создаёт централизованное состояние приложения с явными полями."""
        return {
            "fileType": FILE_TYPE_MAIN,
            "filePaths": [],
            "saveDir": save_dir,
            "routes": [],
            "uniqueProducts": [],
            "filteredRoutes": [],
            "productReplacements": [],
            "routeCategory": ROUTE_CATEGORY_SHK,
            "sortAsc": True,
            "generalFileCreated": False,
            "deptFilesCreated": False,
        }

    def _on_restore_data(self):
        try:
            from ui.pages.restore_data_dialog import run_restore_data_dialog
            if run_restore_data_dialog(self):
                QMessageBox.information(
                    self, "Перезапуск",
                    "Перезапустите приложение для применения восстановленных данных."
                )
                self.close()
        except Exception:
            import logging
            logging.getLogger("app").exception("restore_data")

    def closeEvent(self, event):
        """Проверяет несохранённые изменения перед закрытием."""
        if self._has_unsaved_changes():
            reply = QMessageBox.question(
                self, "Несохранённые изменения",
                "Есть несохранённые изменения (исключённые маршруты или замены продуктов).\n"
                "Сохранить изменения перед закрытием?",
                QMessageBox.StandardButton.Save | QMessageBox.StandardButton.Discard | QMessageBox.StandardButton.Cancel,
                QMessageBox.StandardButton.Cancel,
            )
            if reply == QMessageBox.StandardButton.Cancel:
                event.ignore()
                return
            if reply == QMessageBox.StandardButton.Save:
                self._save_session_before_close()
        event.accept()

    def _has_unsaved_changes(self) -> bool:
        """Проверяет наличие несохранённых изменений (исключённые маршруты, замены продуктов)."""
        routes = self.app_state.get("filteredRoutes") or []
        excluded = sum(1 for r in routes if r.get("excluded"))
        if excluded > 0:
            return True
        replacements = self.app_state.get("productReplacements") or []
        return len(replacements) > 0

    def _save_session_before_close(self) -> None:
        """Сохраняет текущую сессию (маршруты, фильтры) в историю перед закрытием."""
        routes = self.app_state.get("routes") or []
        filtered = self.app_state.get("filteredRoutes") or []
        unique = self.app_state.get("uniqueProducts") or []
        if not routes:
            return
        file_type = self.app_state.get("fileType", FILE_TYPE_MAIN)
        save_dir = self.app_state.get("saveDir") or data_store.get_setting("defaultSaveDir")
        route_category = (routes[0].get("routeCategory", ROUTE_CATEGORY_SHK) if routes else ROUTE_CATEGORY_SHK)
        data_store.save_last_routes(file_type, routes, unique, filtered, route_category=route_category, save_dir=save_dir)

    def _open_file_creation_settings(self):
        try:
            from ui.pages.file_creation_settings_dialog import open_file_creation_settings_dialog
            open_file_creation_settings_dialog(self)
        except Exception:
            import traceback
            import logging
            logging.getLogger("app").exception("file_creation_settings")

    def _open_save_mode_settings(self):
        try:
            from ui.pages.save_mode_settings_dialog import open_save_mode_settings_dialog
            open_save_mode_settings_dialog(self)
        except Exception:
            import traceback
            import logging
            logging.getLogger("app").exception("save_mode_settings")

    def _open_quantity_settings(self):
        try:
            from ui.pages.quantity_settings_dialog import open_quantity_settings_dialog
            open_quantity_settings_dialog(self, self.app_state)
            # После закрытия — обновить preview страницы если открыты
            cb = self.app_state.get("refresh_preview_pages")
            if callable(cb):
                cb()
        except Exception:
            import traceback
            import logging
            logging.getLogger("app").exception("quantity_settings")

    def _open_savings_settings(self):
        try:
            from ui.pages.savings_dialog import open_savings_dialog
            open_savings_dialog(self, self.app_state)
            cb = self.app_state.get("refresh_preview_pages")
            if callable(cb):
                cb()
        except Exception:
            import traceback
            import logging
            logging.getLogger("app").exception("savings_settings")

    def _open_appearance_settings(self):
        try:
            from ui.pages.appearance_settings_dialog import open_appearance_settings_dialog
            open_appearance_settings_dialog(self)
        except Exception:
            import logging
            logging.getLogger("app").exception("appearance_settings")

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
        text = (
            "<h3>Маршруты, Сборка</h3>"
            "<p><b>Инструкция по использованию</b></p>"
            "<p><b>1. Обработка файлов.</b> Вкладка «Обработка файлов»: выберите тип (основной или довоз), "
            "укажите папку сохранения, перетащите XLS-файлы маршрутов (ШК и/или СД) в зону загрузки. "
            "Нажмите «Обработать» — при новых названиях продуктов появится диалог выбора (новый продукт или дубликат).</p>"
            "<p><b>2. Предпросмотр.</b> После обработки откроется таблица общих маршрутов. Используйте поиск (Ctrl+F), "
            "фильтр по продукту, двойной клик по номеру маршрута для редактирования, правый клик — исключить из выгрузки. "
            "Кнопка «Далее» — переход к маршрутам по отделам.</p>"
            "<p><b>3. Маршруты по отделам.</b> Выберите отдел во вкладках, укажите папку сохранения. "
            "«Создать файлы для всех отделов» — генерация XLS-файлов и этикеток. При непривязанных продуктах появится баннер — "
            "откройте «Отделы и продукты» для привязки.</p>"
            "<p><b>4. Этикетки.</b> Выберите отдел и тип, нажмите «Создать файлы этикеток». "
            "В «Настройках этикеток» настройте режимы (чищенка, сыпучка) по отделам.</p>"
            "<p><b>5. Справочники.</b> Меню «Файл» → «Справочники»: «Отделы и продукты» — иерархия отделов и привязка продуктов; "
            "«Продукты» — алиасы (варианты написания → каноническое); «Шаблоны» — структура столбцов XLS по отделам.</p>"
            "<p><b>6. Настройки.</b> «Параметры создания файлов» — размер шрифта и отступы; "
            "«Настройки Количества» — округление в штуках по продуктам и учреждениям.</p>"
            "<p><b>Горячие клавиши:</b> Ctrl+D — Отделы и продукты, Ctrl+T — Шаблоны, Ctrl+P — Продукты, "
            "Ctrl+F — фокус на поиск, Escape — закрыть диалог.</p>"
        )
        QMessageBox.about(self, "О программе", text)

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
        self.app_state.update(self._create_app_state(None))
        self._update_routes_dependent_tabs()
        self.navigate_to.emit("dashboard")


if __name__ == "__main__":
    from PyQt6.QtWidgets import QApplication

    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())
