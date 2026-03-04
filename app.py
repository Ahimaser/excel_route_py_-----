"""
app.py — Точка входа приложения «Маршруты, Сборка».

Изменения:
- Глобальный обработчик исключений (sys.excepthook) — крэш записывается в лог
- Логирование в файл crash.log рядом с app.py
- Ленивая инициализация страниц (создаются при первом переходе)
- Справочники (Отделы, Продукты, Шаблоны) открываются как модальные QDialog.
- Настройки Шт открываются как модальный QDialog.
  После закрытия с «Сохранить» — обновляются превью-страницы.
"""
from __future__ import annotations

import sys
import os
import traceback
import logging

# ─────────────────────────── Логирование ──────────────────────────────────

from core.data_store import get_app_data_dir

_LOG_PATH = str(get_app_data_dir() / "crash.log")

_stream_handler = logging.StreamHandler(sys.stdout)
try:
    import io
    _stream_handler.stream = io.TextIOWrapper(
        sys.stdout.buffer, encoding="utf-8", errors="replace"
    )
except Exception:
    pass
_stream_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(_LOG_PATH, encoding="utf-8"),
        _stream_handler,
    ]
)
log = logging.getLogger("app")


def _global_excepthook(exc_type, exc_value, exc_tb):
    """Перехватывает все необработанные исключения Python и пишет в лог."""
    msg = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
    log.critical("НЕОБРАБОТАННОЕ ИСКЛЮЧЕНИЕ:\n%s", msg)
    try:
        from PyQt6.QtWidgets import QApplication, QMessageBox
        if QApplication.instance():
            QMessageBox.critical(
                None,
                "Критическая ошибка",
                f"Произошла ошибка:\n\n{exc_value}\n\n"
                f"Подробности записаны в файл:\n{_LOG_PATH}"
            )
    except Exception:
        pass
    sys.__excepthook__(exc_type, exc_value, exc_tb)


sys.excepthook = _global_excepthook

# ─────────────────────────── Путь к проекту ───────────────────────────────

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# ─────────────────────────── Qt ───────────────────────────────────────────

from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import Qt

from ui.main_window import MainWindow


# Страницы, которые встраиваются в стек (QWidget-страницы)
PAGE_TITLES = {
    "dashboard":       "Главная",
    "home":            "Обработка файлов",
    "labels":          "Этикетки",
    "preview_general": "Предпросмотр — Общие маршруты",
    "preview_dept":    "Предпросмотр — По отделам",
}

# Справочники — открываются как модальные диалоги (не добавляются в стек)
MODAL_REFS = {"departments", "products", "templates", "settings"}


def main():
    log.info("=== Запуск Маршруты, Сборка ===")

    app = QApplication(sys.argv)
    app.setApplicationName("Маршруты, Сборка")
    app.setOrganizationName("RouteManager")

    try:
        from PyQt6.QtCore import qInstallMessageHandler
        def _qt_msg(mode, context, message):
            if "error" in message.lower() or "warning" in message.lower():
                log.warning("Qt: %s", message)
        qInstallMessageHandler(_qt_msg)
    except Exception:
        pass

    try:
        window = MainWindow()
    except Exception:
        log.critical("Ошибка при создании MainWindow:\n%s", traceback.format_exc())
        raise

    stack = window.stack

    _page_cache: dict[str, object] = {}
    _page_idx:   dict[str, int]    = {}

    # ── Модальные диалоги справочников ────────────────────────────────────

    def _open_departments():
        """Открывает модальный диалог «Отделы и продукты»."""
        log.debug("Открываем диалог: departments")
        try:
            from ui.pages.departments_page import open_modal
            open_modal(window, window.app_state)
            # После закрытия — обновить preview_dept если открыта
            pd = _page_cache.get("preview_dept")
            if pd is not None and hasattr(pd, "refresh"):
                pd.refresh()
        except Exception:
            log.critical("Ошибка при открытии departments:\n%s", traceback.format_exc())

    def _open_products():
        """Открывает модальный диалог «Справочник продуктов»."""
        log.debug("Открываем диалог: products")
        try:
            from ui.pages.products_page import open_modal
            open_modal(window, window.app_state)
            if window.app_state.get("open_departments_after_products"):
                window.app_state["open_departments_after_products"] = False
                _open_departments()
        except Exception:
            log.critical("Ошибка при открытии products:\n%s", traceback.format_exc())

    def _open_templates():
        """Открывает модальный диалог «Шаблоны»."""
        log.debug("Открываем диалог: templates")
        try:
            from ui.pages.templates_page import open_modal
            open_modal(window, window.app_state)
        except Exception:
            log.critical("Ошибка при открытии templates:\n%s", traceback.format_exc())

    def _open_settings():
        """Открывает модальный диалог «Настройки Шт»."""
        log.debug("Открываем диалог: settings")
        try:
            from ui.pages.settings_dialog import open_settings_dialog

            def _on_saved():
                log.debug("settings: on_saved — обновляем превью-страницы")
                pg = _page_cache.get("preview_general")
                if pg is not None and hasattr(pg, "_render_table"):
                    pg._render_table()
                pd = _page_cache.get("preview_dept")
                if pd is not None and hasattr(pd, "refresh"):
                    pd.refresh()

            open_settings_dialog(window, on_saved=_on_saved)
        except Exception:
            log.critical("Ошибка при открытии settings:\n%s", traceback.format_exc())

    # Словарь: имя → функция открытия модального диалога
    _modal_openers = {
        "departments": _open_departments,
        "products":    _open_products,
        "templates":   _open_templates,
        "settings":    _open_settings,
    }

    # ── Очистка маршрутов ───────────────────────────────────────────────────

    def _clear_last_routes_only():
        """Очищает только сохранённые «последние» маршруты (остаётся на дашборде)."""
        from core import data_store
        data_store.clear_last_routes()
        try:
            from PyQt6.QtWidgets import QMessageBox
            QMessageBox.information(window, "Готово", "Сохранённые маршруты удалены.")
        except Exception:
            pass

    def _clear_routes_and_go_dashboard():
        """Очищает app_state и последние маршруты, переходит на дашборд."""
        from core import data_store
        data_store.clear_last_routes()
        window.app_state.update({
            "filePaths": [], "routes": [], "uniqueProducts": [],
            "filteredRoutes": [], "routeCategory": "ШК",
        })
        home = _page_cache.get("home")
        if home and hasattr(home, "reset"):
            home.reset()
        navigate("dashboard")

    # ── Загрузка последних маршрутов и переход в превью ───────────────────

    def _load_last_and_go_preview(file_type: str):
        """Загружает последние маршруты (main/increase) в app_state и переходит в preview_general."""
        from core import data_store
        data = data_store.get_last_routes(file_type)
        if not data:
            from PyQt6.QtWidgets import QMessageBox
            QMessageBox.warning(
                window, "Нет данных",
                "Нет сохранённых маршрутов для выбранного типа. Сначала обработайте файлы."
            )
            return
        window.app_state.update({
            "fileType":       file_type,
            "routes":         data.get("routes", []),
            "uniqueProducts": data.get("uniqueProducts", []),
            "filteredRoutes": data.get("filteredRoutes", []),
            "routeCategory":  data.get("routeCategory") or "ШК",
        })
        navigate("preview_general")

    # ── Страницы в стеке ──────────────────────────────────────────────────

    def _get_page(name: str):
        """Возвращает страницу (создаёт при первом обращении)."""
        if name in _page_cache:
            return _page_cache[name]

        log.debug("Создаём страницу: %s", name)
        try:
            if name == "dashboard":
                from ui.pages.dashboard_page import DashboardPage
                page = DashboardPage(window.app_state)
                page.go_process_files.connect(lambda: navigate("home"))
                page.open_last_main.connect(lambda: _load_last_and_go_preview("main"))
                page.open_last_increase.connect(lambda: _load_last_and_go_preview("increase"))
                page.go_labels.connect(lambda: navigate("labels"))
                page.clear_last_routes.connect(_clear_last_routes_only)

            elif name == "home":
                from ui.pages.home_page import HomePage
                page = HomePage(window.app_state)
                page.go_preview.connect(lambda: navigate("preview_general"))

            elif name == "labels":
                from ui.pages.labels_page import LabelsPage
                page = LabelsPage(window.app_state)
                page.go_back.connect(lambda: navigate("dashboard"))
                page.go_open_routes.connect(lambda: _load_last_and_go_preview("main"))
                page.go_process_files.connect(lambda: navigate("home"))

            elif name == "preview_general":
                from ui.pages.preview_general_page import PreviewGeneralPage
                page = PreviewGeneralPage(window.app_state)
                page.go_back.connect(lambda: navigate("home"))
                page.go_dept_preview.connect(lambda: navigate("preview_dept"))
                page.go_settings.connect(_open_settings)
                page.go_clear_routes.connect(_clear_routes_and_go_dashboard)

            elif name == "preview_dept":
                from ui.pages.preview_dept_page import PreviewDeptPage
                page = PreviewDeptPage(window.app_state)
                page.go_back.connect(lambda: navigate("preview_general"))
                page.go_clear_routes.connect(_clear_routes_and_go_dashboard)

            else:
                log.warning("Неизвестная страница: %s", name)
                return None

        except Exception:
            log.critical("Ошибка при создании страницы '%s':\n%s",
                         name, traceback.format_exc())
            raise

        idx = stack.addWidget(page)
        _page_cache[name] = page
        _page_idx[name]   = idx
        log.debug("Страница '%s' создана, idx=%d", name, idx)
        return page

    # ── Навигация ─────────────────────────────────────────────────────────

    def navigate(page_name: str):
        """Переходит на страницу стека или открывает модальный диалог."""
        log.debug("navigate -> %s", page_name)

        # Справочники — модальные диалоги
        if page_name in _modal_openers:
            _modal_openers[page_name]()
            return

        # Обычные страницы стека
        if page_name not in PAGE_TITLES:
            log.debug("navigate: пропускаем неизвестное имя '%s'", page_name)
            return

        try:
            page = _get_page(page_name)
            if page is None:
                return
            idx = _page_idx[page_name]
            stack.setCurrentIndex(idx)
            window.set_page_title(PAGE_TITLES.get(page_name, ""))
            if hasattr(page, "refresh"):
                page.refresh()
        except Exception:
            log.critical("Ошибка при переходе на страницу '%s':\n%s",
                         page_name, traceback.format_exc())
            try:
                from PyQt6.QtWidgets import QMessageBox
                QMessageBox.warning(
                    window, "Ошибка",
                    f"Не удалось открыть страницу «{PAGE_TITLES.get(page_name, page_name)}».\n\n"
                    f"Подробности в файле:\n{_LOG_PATH}"
                )
            except Exception:
                pass

    window.navigate_to.connect(navigate)

    # ── Новая сессия ──────────────────────────────────────────────────────

    def new_session():
        window.app_state.update({
            "fileType":       "main",
            "filePaths":      [],
            "saveDir":        None,
            "routes":         [],
            "uniqueProducts": [],
            "filteredRoutes": [],
            "routeCategory":  "ШК",
            "sortAsc":        False,
        })
        home = _page_cache.get("home")
        if home and hasattr(home, "reset"):
            home.reset()
        navigate("dashboard")

    window._on_new_session = new_session

    log.info("Переход на стартовую страницу")
    navigate("dashboard")
    window.show()
    log.info("Окно показано, запускаем event loop")

    exit_code = app.exec()
    log.info("Приложение завершено с кодом %d", exit_code)
    sys.exit(exit_code)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        log.critical("Фатальная ошибка в main():\n%s", traceback.format_exc())
        sys.exit(1)
