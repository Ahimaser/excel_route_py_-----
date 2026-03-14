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
from PyQt6.QtGui import QIcon

from ui.main_window import MainWindow


# Страницы, которые встраиваются в стек (QWidget-страницы)
PAGE_TITLES = {
    "dashboard":       "Главная",
    "home":            "Обработка файлов",
    "labels":          "Этикетки",
    "preview_general": "Предпросмотр — Общие маршруты",
    "preview_dept":    "Предпросмотр — По отделам",
}

# Краткие подсказки (при наведении на «!»)
PAGE_HINTS_SHORT = {
    "dashboard": "Главная: обработка файлов, последние маршруты, этикетки.",
    "home": "Загрузка XLS, выбор папки, обработка.",
    "labels": "Этикетки по шаблонам продуктов.",
    "preview_general": "Таблица маршрутов, поиск, фильтр, создание файла.",
    "preview_dept": "Маршруты по отделам, фильтр, генерация файлов.",
}

# Подробные инструкции по пунктам (при нажатии на «!»)
PAGE_HINTS_LONG = {
    "dashboard": (
        "Инструкция — Главная страница\n\n"
        "1. «Обработать файлы» — переход к загрузке XLS-файлов маршрутов (ШК и/или СД).\n"
        "2. «Последние (основной)» — открыть последние сохранённые маршруты основного типа.\n"
        "3. «Последние (довоз)» — открыть последние сохранённые маршруты довоза.\n"
        "4. «Этикетки» — переход к созданию этикеток XLS по шаблонам.\n"
        "5. «Очистить» — удалить сохранённые «последние» маршруты из памяти."
    ),
    "home": (
        "Инструкция — Обработка файлов\n\n"
        "1. Выберите тип файла: основной или довоз (увеличение).\n"
        "2. Укажите папку сохранения результатов.\n"
        "3. Перетащите XLS-файлы в зону загрузки или нажмите «Выбрать файлы».\n"
        "4. При необходимости измените категорию маршрутов (ШК/СД) для округления.\n"
        "5. Нажмите «Обработать» — после обработки откроется предпросмотр."
    ),
    "labels": (
        "Инструкция — Этикетки\n\n"
        "1. Убедитесь, что маршруты загружены (обработайте файлы или откройте последние).\n"
        "2. Выберите продукт в списке и нажмите «Загрузить шаблон» — укажите XLS-шаблон для этого продукта.\n"
        "3. Либо откройте «Настройки этикеток» и задайте шаблон для продуктов по отделам.\n"
        "4. Нажмите «Создать XLS по шаблонам» — файлы сохранятся в папку «Этикетки на ДД.ММ.ГГГГ»."
    ),
    "preview_general": (
        "Инструкция — Предпросмотр (общие маршруты)\n\n"
        "1. Поиск: введите текст в поле поиска или нажмите Ctrl+F.\n"
        "2. Фильтр по продукту: выберите продукт в выпадающем списке.\n"
        "3. Двойной клик по номеру маршрута — изменить номер.\n"
        "4. Правый клик по строке — исключить маршрут из выгрузки или удалить.\n"
        "5. Ctrl+колёсико мыши над таблицей — изменить размер шрифта.\n"
        "6. «Создать файл» — сформировать Excel «Общие маршруты».\n"
        "7. «Этикетки» — создать этикетки по шаблонам.\n"
        "8. «Справочник продуктов» — настройки продуктов и кол-во в шт (ПКМ по продукту)."
    ),
    "preview_dept": (
        "Инструкция — Маршруты по отделам\n\n"
        "1. В фильтре «Показать отдел/подотдел» выберите нужный отдел или «Все отделы».\n"
        "2. Вкладки — по одной на каждый отдел/подотдел с таблицей маршрутов.\n"
        "3. Ctrl+колёсико мыши над таблицей — изменить размер шрифта.\n"
        "4. «Сгенерировать все» — создать Excel-файлы по отделам в выбранную папку.\n"
        "5. «Этикетки» — создать этикетки в папку «Этикетки на ДД.ММ.ГГГГ».\n"
        "6. «Отделы и продукты» — открыть настройку привязки продуктов к отделам."
    ),
}

# Справочники — открываются как модальные диалоги (не добавляются в стек)
MODAL_REFS = {"departments", "products", "templates"}


def main():
    log.info("=== Запуск Маршруты, Сборка ===")

    app = QApplication(sys.argv)
    app.setApplicationName("Маршруты, Сборка")
    app.setOrganizationName("RouteManager")
    app_icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "app_icon.svg")
    if os.path.isfile(app_icon_path):
        app_icon = QIcon(app_icon_path)
        if not app_icon.isNull():
            app.setWindowIcon(app_icon)

    # Тема в стиле веб-приложения (Material Design)
    try:
        from qt_material import apply_stylesheet
        apply_stylesheet(app, theme="light_blue.xml", invert_secondary=True)
        try:
            from ui.styles import QUANTITY_DIALOG_EXTRA, RIBBON_TABS_EXTRA
            app.setStyleSheet(app.styleSheet() + RIBBON_TABS_EXTRA + QUANTITY_DIALOG_EXTRA)
        except ImportError:
            pass
    except ImportError:
        from ui.styles import STYLESHEET
        app.setStyleSheet(STYLESHEET)

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
    if os.path.isfile(app_icon_path):
        app_icon = QIcon(app_icon_path)
        if not app_icon.isNull():
            window.setWindowIcon(app_icon)

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
            # После закрытия — обновить preview страницы если открыты
            for name in ("preview_dept", "preview_general"):
                p = _page_cache.get(name)
                if p is not None and hasattr(p, "refresh"):
                    p.refresh()
        except Exception:
            log.critical("Ошибка при открытии departments:\n%s", traceback.format_exc())
            try:
                from ui.widgets import message_plain
                from PyQt6.QtWidgets import QMessageBox
                message_plain(
                    window, "Ошибка",
                    "Не удалось открыть диалог «Отделы и продукты».\n\nПодробности в файле:\n" + _LOG_PATH,
                    icon=QMessageBox.Icon.Warning,
                )
            except Exception:
                pass

    def _open_products():
        """Открывает модальный диалог «Справочник продуктов»."""
        log.debug("Открываем диалог: products")
        try:
            from ui.pages.products_page import open_modal
            open_modal(window, window.app_state)
            if window.app_state.get("open_departments_after_products"):
                window.app_state["open_departments_after_products"] = False
                _open_departments()
            for name in ("preview_dept", "preview_general"):
                p = _page_cache.get(name)
                if p is not None and hasattr(p, "refresh"):
                    p.refresh()
        except Exception:
            log.critical("Ошибка при открытии products:\n%s", traceback.format_exc())
            try:
                from ui.widgets import message_plain
                from PyQt6.QtWidgets import QMessageBox
                message_plain(
                    window, "Ошибка",
                    "Не удалось открыть диалог «Справочник продуктов».\n\nПодробности в файле:\n" + _LOG_PATH,
                    icon=QMessageBox.Icon.Warning,
                )
            except Exception:
                pass

    def _open_templates():
        """Открывает модальный диалог «Шаблоны»."""
        log.debug("Открываем диалог: templates")
        try:
            from ui.pages.templates_page import open_modal
            open_modal(window, window.app_state)
        except Exception:
            log.critical("Ошибка при открытии templates:\n%s", traceback.format_exc())
            try:
                from ui.widgets import message_plain
                from PyQt6.QtWidgets import QMessageBox
                message_plain(
                    window, "Ошибка",
                    "Не удалось открыть диалог «Шаблоны».\n\nПодробности в файле:\n" + _LOG_PATH,
                    icon=QMessageBox.Icon.Warning,
                )
            except Exception:
                pass

    def _open_quantity_settings():
        """Открывает модальный диалог «Настройки Количества»."""
        log.debug("Открываем диалог: quantity_settings")
        try:
            from ui.pages.quantity_settings_dialog import open_quantity_settings_dialog
            open_quantity_settings_dialog(window, window.app_state)
        except Exception:
            log.critical("Ошибка при открытии Настройки Количества:\n%s", traceback.format_exc())
            try:
                from ui.widgets import message_plain
                from PyQt6.QtWidgets import QMessageBox
                message_plain(
                    window,
                    "Ошибка",
                    "Не удалось открыть «Настройки Количества».\n\nПодробности в файле:\n" + _LOG_PATH,
                    icon=QMessageBox.Icon.Warning,
                )
            except Exception:
                pass

    # Словарь: имя → функция открытия модального диалога
    _modal_openers = {
        "departments": _open_departments,
        "products":    _open_products,
        "templates":   _open_templates,
    }

    # ── Очистка маршрутов ───────────────────────────────────────────────────

    def _clear_routes_and_go_dashboard():
        """Очищает app_state и последние маршруты, переходит на дашборд."""
        from core import data_store
        data_store.clear_last_routes()
        window.app_state.update({
            "fileType": "main", "filePaths": [], "routes": [],
            "uniqueProducts": [], "filteredRoutes": [],
            "routeCategory": "ШК", "sortAsc": True,
            "institutionList": [],
        })
        home = _page_cache.get("home")
        if home and hasattr(home, "reset"):
            home.reset()
        set_status = window.app_state.get("set_status")
        if callable(set_status):
            set_status("Маршруты очищены")
        navigate("dashboard")

    # ── Загрузка маршрутов и переход в превью ─────────────────────────────

    def _apply_saved_blob_and_open_preview(data: dict, fallback_file_type: str) -> None:
        """Применяет данные сохранения к app_state и переходит в preview_general."""
        import copy
        from core import data_store
        routes = copy.deepcopy(data.get("routes") or [])
        unique_products = copy.deepcopy(data.get("uniqueProducts") or [])
        filtered = copy.deepcopy(data.get("filteredRoutes") or data.get("routes") or [])
        file_type = data.get("fileType") or fallback_file_type
        n = len(filtered)
        set_status = window.app_state.get("set_status")
        if callable(set_status) and n:
            set_status(f"Загружено {n} маршрутов")
        window.app_state["institutionList"] = data_store.get_institution_list_from_routes(filtered)
        window.app_state.update({
            "fileType":       file_type,
            "routes":         routes,
            "uniqueProducts": unique_products,
            "filteredRoutes": filtered,
            "routeCategory":  data.get("routeCategory") or "ШК",
        })
        navigate("preview_general")

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
        _apply_saved_blob_and_open_preview(data, file_type)

    def _open_history_and_go_preview(file_type: str) -> None:
        """Открывает диалог «История маршрутов», выбирает запись и переходит в preview."""
        from ui.pages.routes_history_dialog import pick_routes_history_entry
        data = pick_routes_history_entry(window, file_type=file_type)
        if not data:
            return
        _apply_saved_blob_and_open_preview(data, file_type)

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
                page.open_history.connect(lambda: _open_history_and_go_preview(None))
                page.go_process_files.connect(lambda: navigate("home"))
                page.go_last_main.connect(lambda: _load_last_and_go_preview("main"))
                page.go_last_increase.connect(lambda: _load_last_and_go_preview("increase"))
                page.go_labels.connect(lambda: navigate("labels"))
                page.go_clear.connect(_clear_routes_and_go_dashboard)

            elif name == "home":
                from ui.pages.home_page import HomePage
                page = HomePage(window.app_state)
                page.go_preview.connect(lambda: navigate("preview_general"))

            elif name == "labels":
                from ui.pages.labels_page import LabelsPage
                page = LabelsPage(window.app_state)
                page.go_back.connect(lambda: navigate("dashboard"))
                page.go_open_routes.connect(lambda: _open_history_and_go_preview("main"))
                page.go_process_files.connect(lambda: navigate("home"))

            elif name == "preview_general":
                from ui.pages.preview_general_page import PreviewGeneralPage
                page = PreviewGeneralPage(window.app_state)
                page.go_back.connect(lambda: navigate("home"))
                page.go_home.connect(lambda: navigate("dashboard"))
                page.go_dept_preview.connect(lambda: navigate("preview_dept"))
                page.go_settings.connect(_open_products)
                page.go_clear_routes.connect(_clear_routes_and_go_dashboard)

            elif name == "preview_dept":
                from ui.pages.preview_dept_page import PreviewDeptPage
                page = PreviewDeptPage(window.app_state)
                page.go_back.connect(lambda: navigate("preview_general"))
                page.go_home.connect(lambda: navigate("dashboard"))
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
            window.set_ribbon_page(page_name)
            window.set_page_title(PAGE_TITLES.get(page_name, ""))
            window.set_page_hint(
                PAGE_HINTS_SHORT.get(page_name, ""),
                PAGE_HINTS_LONG.get(page_name, ""),
            )
            if hasattr(page, "refresh"):
                page.refresh()
            if hasattr(window, "_update_routes_dependent_tabs"):
                window._update_routes_dependent_tabs()
        except Exception:
            log.critical("Ошибка при переходе на страницу '%s':\n%s",
                         page_name, traceback.format_exc())
            try:
                from ui.widgets import message_plain
                from PyQt6.QtWidgets import QMessageBox
                message_plain(
                    window, "Ошибка",
                    f"Не удалось открыть страницу «{PAGE_TITLES.get(page_name, page_name)}».\n\n"
                    f"Подробности в файле:\n{_LOG_PATH}",
                    icon=QMessageBox.Icon.Warning,
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
            "sortAsc":        True,
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
