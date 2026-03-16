"""
preview_dept_page.py — Предпросмотр и генерация файлов по отделам/подотделам.

Оптимизации:
- Используем data_store.get_products_map() для быстрого доступа к настройкам продуктов
- Таблицы строятся с setUpdatesEnabled(False) для batch-рендера
- _build_dept_groups() использует индекс продуктов {deptKey: [routes]}
- Worker получает prod_map вместо products_settings (без лишних преобразований)
"""
from __future__ import annotations

import os
from datetime import date, datetime, timedelta
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFrame, QTableWidget, QTableWidgetItem, QHeaderView,
    QAbstractItemView, QMessageBox, QFileDialog, QComboBox,
    QProgressBar, QApplication, QScrollArea, QStackedWidget,
    QCheckBox, QSizePolicy, QTabWidget,
)
from PyQt6.QtCore import Qt, pyqtSignal, QThread, QObject, QEvent, QTimer
from PyQt6.QtGui import QFont, QBrush, QColor, QWheelEvent

from core import data_store, excel_generator
from ui.pages import departments_page as dept_mod


# ─────────────────────────── Worker ───────────────────────────────────────

class DeptGenerateWorker(QObject):
    finished = pyqtSignal(list)
    error    = pyqtSignal(str)

    def __init__(self, dept_groups: list, file_type: str, save_dir: str,
                 prod_map: dict, templates: list, sort_asc: bool = True,
                 all_routes: list | None = None, general_path: str | None = None,
                 date_str: str | None = None, replacements: list | None = None):
        super().__init__()
        self.dept_groups = dept_groups
        self.file_type   = file_type
        self.save_dir    = save_dir
        self.prod_map    = prod_map
        self.templates   = templates
        self.sort_asc    = sort_asc
        self.all_routes  = all_routes or []
        self.general_path = general_path
        self.date_str = date_str
        self.replacements = replacements or []

    def run(self):
        try:
            replaced_routes = excel_generator.apply_replacements(
                self.all_routes, self.replacements, self.sort_asc
            )
            dept_groups = data_store.build_dept_groups_from_routes(replaced_routes)
            created: list[str] = []
            if self.file_type == "main":
                for cat in ("ШК", "СД"):
                    cat_routes = [r for r in replaced_routes if (r.get("routeCategory") or "ШК") == cat]
                    if not cat_routes:
                        continue
                    path = excel_generator.get_general_routes_path(
                        self.save_dir, self.file_type, self.date_str, route_category=cat
                    )
                    os.makedirs(os.path.dirname(path), exist_ok=True)
                    excel_generator.generate_general_routes(
                        cat_routes,
                        self.file_type,
                        path,
                        self.prod_map,
                        self.sort_asc,
                        date_str=self.date_str,
                        replacements=[],
                    )
                    created.append(path)
            elif self.general_path:
                os.makedirs(os.path.dirname(self.general_path), exist_ok=True)
                excel_generator.generate_general_routes(
                    replaced_routes,
                    self.file_type,
                    self.general_path,
                    self.prod_map,
                    self.sort_asc,
                    date_str=self.date_str,
                    replacements=[],
                )
                created.append(self.general_path)
            departments_ref = data_store.get_ref("departments") or []
            dept_created = excel_generator.generate_dept_files(
                dept_groups, self.file_type,
                self.save_dir, self.prod_map, self.templates,
                self.sort_asc, date_str=self.date_str,
                replacements=self.replacements,
                departments_ref=departments_ref,
            )
            created.extend(dept_created)
            day_dir = excel_generator.get_routes_day_folder(self.save_dir, self.date_str)
            products_ref = data_store.get_ref("products") or []
            report_paths = excel_generator.generate_pcs_compare_report(
                day_dir,
                main_routes=replaced_routes if self.file_type == "main" else [],
                increase_routes=replaced_routes if self.file_type == "increase" else [],
                products_ref=products_ref,
                date_str=self.date_str,
            )
            created.extend(report_paths)
            self.finished.emit(created)
        except Exception as e:
            self.error.emit(str(e))


# ─────────────────────────── Страница ─────────────────────────────────────

class PreviewDeptPage(QWidget):
    """Предпросмотр и генерация файлов по отделам."""

    go_back = pyqtSignal()
    go_home = pyqtSignal()  # Переход на главную (dashboard)
    go_clear_routes = pyqtSignal()
    go_open_last = pyqtSignal()
    go_process_files = pyqtSignal()

    def __init__(self, app_state: dict):
        super().__init__()
        self.app_state   = app_state
        self._dept_groups: list[dict] = []
        self._table_font_size = 11
        self._current_group_idx = -1
        self._dept_tab_defs: list[dict] = []
        self._dept_buttons: list[QPushButton] = []
        self._dept_tables: list[QTableWidget] = []
        self._refreshing = False
        self._build_ui()

    def _build_ui(self):
        main_lay = QVBoxLayout(self)
        main_lay.setContentsMargins(0, 0, 0, 0)
        top = QWidget()
        top.setObjectName("previewDeptContent")
        inner = QVBoxLayout(top)
        inner.setContentsMargins(20, 16, 20, 16)
        inner.setSpacing(12)

        h_row = QHBoxLayout()
        self.lbl_title = QLabel("Маршруты по отделам")
        self.lbl_title.setObjectName("sectionTitle")
        h_row.addWidget(self.lbl_title)
        h_row.addStretch()

        inner.addLayout(h_row)

        # Баннер непривязанных продуктов (4A)
        self.banner_unassigned = QFrame()
        self.banner_unassigned.setObjectName("bannerWarning")
        self.banner_unassigned.setVisible(False)
        banner_lay = QHBoxLayout(self.banner_unassigned)
        banner_lay.setContentsMargins(12, 8, 12, 8)
        self.lbl_banner = QLabel("")
        self.lbl_banner.setWordWrap(True)
        banner_lay.addWidget(self.lbl_banner)
        self.btn_banner_depts = QPushButton("Открыть Отделы и продукты")
        self.btn_banner_depts.setObjectName("btnPrimary")
        self.btn_banner_depts.clicked.connect(self._on_banner_open_departments)
        banner_lay.addWidget(self.btn_banner_depts)
        self.btn_banner_products = QPushButton("Справочник продуктов")
        self.btn_banner_products.setObjectName("btnSecondary")
        self.btn_banner_products.clicked.connect(self._on_banner_open_products)
        banner_lay.addWidget(self.btn_banner_products)
        inner.addWidget(self.banner_unassigned)

        # Выбор отдела/подотдела: отделы сверху, подотделы под ними (скрыты до выбора)
        dept_row = QHBoxLayout()
        dept_row.setSpacing(8)
        dept_tabs_frame = QFrame()
        dept_tabs_frame.setObjectName("deptTabsBar")
        self.dept_btns_lay = QHBoxLayout(dept_tabs_frame)
        self.dept_btns_lay.setContentsMargins(8, 6, 8, 6)
        self.dept_btns_lay.setSpacing(6)
        dept_row.addWidget(dept_tabs_frame)
        dept_row.addStretch()
        inner.addLayout(dept_row)

        self._content_stack = QStackedWidget()
        self._empty_widget = QWidget()
        empty_lay = QVBoxLayout(self._empty_widget)
        empty_lay.setAlignment(Qt.AlignmentFlag.AlignCenter)
        empty_lay.setSpacing(16)
        lbl_empty = QLabel("Нет данных для отображения")
        lbl_empty.setObjectName("cardTitle")
        lbl_empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
        empty_lay.addWidget(lbl_empty)
        hint_empty = QLabel(
            "Загрузите XLS файлы или создайте отделы с привязанными продуктами."
        )
        hint_empty.setObjectName("stepLabel")
        hint_empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
        hint_empty.setWordWrap(True)
        empty_lay.addWidget(hint_empty)
        btn_row_empty = QHBoxLayout()
        btn_row_empty.setSpacing(10)
        btn_open_last = QPushButton("Открыть последние")
        btn_open_last.setObjectName("btnPrimary")
        btn_open_last.clicked.connect(self._on_empty_open_last)
        btn_process = QPushButton("Обработать файлы")
        btn_process.setObjectName("btnSecondary")
        btn_process.clicked.connect(self._on_empty_process_files)
        btn_row_empty.addWidget(btn_open_last)
        btn_row_empty.addWidget(btn_process)
        empty_lay.addLayout(btn_row_empty)
        self._content_stack.addWidget(self._empty_widget)
        self._single_table_container = QWidget()
        self._tbl_lay = QVBoxLayout(self._single_table_container)
        self._tbl_lay.setContentsMargins(6, 6, 6, 6)
        self._tbl_lay.setSpacing(6)
        self._dept_hint_label = QLabel("")
        self._dept_hint_label.setObjectName("hintLabel")
        self._tbl_lay.addWidget(self._dept_hint_label)
        self._dept_table = None
        self._dept_save_btn = QPushButton("")
        self._dept_save_btn.setObjectName("btnSecondary")
        self._dept_save_btn.clicked.connect(self._on_save_current_dept)
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_row.addWidget(self._dept_save_btn)
        self._tbl_lay.addLayout(btn_row)
        self._content_stack.addWidget(self._single_table_container)
        self._subdept_tabs_widget = QTabWidget()
        self._subdept_tabs_widget.setDocumentMode(True)
        self._content_stack.addWidget(self._subdept_tabs_widget)
        self._content_stack.setMinimumHeight(400)
        main_lay.addWidget(top)
        main_lay.addWidget(self._content_stack, 1)

        # Нижняя панель с кнопкой создания файлов — всегда видна (вне прокрутки)
        bottom_row = QHBoxLayout()
        bottom_row.addWidget(QLabel("Папка сохранения:"))
        self.lbl_save_dir = QLabel("")
        self.lbl_save_dir.setObjectName("hintLabel")
        self.lbl_save_dir.setMinimumWidth(200)
        self.lbl_save_dir.setWordWrap(True)
        bottom_row.addWidget(self.lbl_save_dir, 1)
        self.btn_choose_save_dir = QPushButton("Выбрать папку")
        self.btn_choose_save_dir.setObjectName("btnSecondary")
        self.btn_choose_save_dir.clicked.connect(self._on_choose_save_dir)
        bottom_row.addWidget(self.btn_choose_save_dir)
        self.btn_save_mode = QPushButton("Режимы сохранения (ШК/СД)")
        self.btn_save_mode.setObjectName("btnSecondary")
        self.btn_save_mode.setToolTip(
            "Настроить режим сохранения по отделам для школ и садов: "
            "все в один файл, по группам или по продуктам."
        )
        self.btn_save_mode.clicked.connect(self._on_open_save_mode_settings)
        bottom_row.addWidget(self.btn_save_mode)
        bottom_row.addStretch()
        self.btn_generate_all = QPushButton("Создать файлы для всех отделов")
        self.btn_generate_all.setMinimumWidth(220)
        self.btn_generate_all.setObjectName("btnPrimary")
        self.btn_generate_all.clicked.connect(self._on_generate_all)
        bottom_row.addWidget(self.btn_generate_all)

        main_lay.addLayout(bottom_row)

        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.setVisible(False)
        main_lay.addWidget(self.progress)

    # ─────────────────────────── Проверка отделов ────────────────────────

    def _check_unassigned_products(self) -> list[str]:
        """Возвращает список продуктов без отдела из текущих маршрутов."""
        routes   = self.app_state.get("filteredRoutes", [])
        products = data_store.get_ref("products") or []
        aliases  = data_store.get_aliases()
        assigned = {p["name"] for p in products if p.get("deptKey")}
        unassigned: set[str] = set()
        for r in routes:
            if r.get("excluded"):
                continue
            for p in r.get("products", []):
                canonical = aliases.get(p["name"], p["name"])
                if canonical not in assigned:
                    unassigned.add(p["name"])
        return sorted(unassigned)

    # ─────────────────────────── Данные ───────────────────────────────────

    def _build_dept_groups(self) -> list[dict]:
        """
        Строит список групп {key, name, routes} для каждого отдела/подотдела.
        Использует индекс {deptKey: [route_indices]} для O(n) вместо O(n*m).
        """
        routes   = self.app_state.get("filteredRoutes", [])
        depts    = data_store.get_ref("departments") or []
        products = data_store.get_ref("products") or []

        # Индекс: deptKey -> список продуктов
        prod_by_dept: dict[str, list[str]] = {}
        for p in products:
            k = p.get("deptKey")
            if k:
                prod_by_dept.setdefault(k, []).append(p["name"])

        # Для каждого отдела/подотдела собираем маршруты, в которых есть его продукты
        groups: list[dict] = []

        def _aggregate_products(prods: list[dict]) -> list[dict]:
            """Объединяет дубликаты продуктов: один продукт — одна строка, количества суммируются."""
            by_name: dict[str, dict] = {}
            for p in prods:
                name = p.get("name", "")
                if not name:
                    continue
                if name in by_name:
                    agg = by_name[name]
                    try:
                        q = float(agg.get("quantity") or 0) + float(p.get("quantity") or 0)
                        agg["quantity"] = q
                    except (TypeError, ValueError):
                        pass
                else:
                    by_name[name] = dict(p)
            return list(by_name.values())

        def _collect_routes(dept_key: str) -> list[dict]:
            prod_names = set(prod_by_dept.get(dept_key, []))
            if not prod_names:
                return []
            result = []
            for r in routes:
                if r.get("excluded"):
                    continue
                dept_prods = [
                    p for p in r.get("products", [])
                    if p["name"] in prod_names
                ]
                if dept_prods:
                    result.append({
                        "routeNum":       r["routeNum"],
                        "address":        r["address"],
                        "routeCategory":  r.get("routeCategory") or "ШК",
                        "products":       _aggregate_products(dept_prods),
                    })
            return result

        for dept in depts:
            # Подотделы
            for sub in dept.get("subdepts", []):
                sub_routes = _collect_routes(sub["key"])
                if sub_routes:
                    groups.append({
                        "key":              sub["key"],
                        "name":             sub["name"],
                        "is_subdept":       True,
                        "parent_dept_name": dept.get("name") or dept.get("key", ""),
                        "routes":           sub_routes,
                    })

            # Сам отдел (продукты, привязанные напрямую к отделу)
            dept_routes = _collect_routes(dept["key"])
            if dept_routes:
                groups.append({
                    "key":        dept["key"],
                    "name":       dept["name"],
                    "is_subdept": False,
                    "routes":     dept_routes,
                })

        return groups

    def _build_dept_tab_defs(self) -> list[dict]:
        """Строит структуру: отделы с подотделами (как в настройках округления Шт)."""
        depts = data_store.get_ref("departments") or []
        result: list[dict] = []
        for dept in depts:
            dept_key = dept.get("key") or ""
            dept_name = dept.get("name") or dept_key
            if not dept_key:
                continue
            scopes: list[tuple[str, int]] = []
            for sub in dept.get("subdepts", []):
                sub_key = sub.get("key") or ""
                sub_name = sub.get("name") or sub_key
                for i, g in enumerate(self._dept_groups):
                    if g.get("is_subdept") and g.get("key") == sub_key:
                        scopes.append((f"{sub_name} ({len(g['routes'])})", i))
                        break
            for i, g in enumerate(self._dept_groups):
                if not g.get("is_subdept") and g.get("key") == dept_key:
                    scopes.append((f"{dept_name} — отдел ({len(g['routes'])})", i))
                    break
            if scopes:
                total = sum(len(self._dept_groups[idx]["routes"]) for _, idx in scopes)
                result.append({
                    "dept_key": dept_key,
                    "dept_name": dept_name,
                    "tab_text": f"{dept_name} ({total})",
                    "scopes": scopes,
                })
        return result

    def _clear_dept_buttons(self) -> None:
        while self.dept_btns_lay.count():
            item = self.dept_btns_lay.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
        self._dept_buttons.clear()

    def _clear_subdept_tabs(self) -> None:
        while self._subdept_tabs_widget.count():
            w = self._subdept_tabs_widget.widget(0)
            self._subdept_tabs_widget.removeTab(0)
            if w:
                w.deleteLater()
        self._dept_tables.clear()

    def _populate_dept_buttons(self) -> None:
        for i, tab_def in enumerate(self._dept_tab_defs):
            btn = QPushButton(tab_def["tab_text"])
            btn.setObjectName("deptTab")
            btn.setCheckable(True)
            btn.setChecked(i == 0)
            btn.clicked.connect(lambda checked, idx=i: self._on_dept_clicked(idx))
            self.dept_btns_lay.addWidget(btn)
            self._dept_buttons.append(btn)
        self.dept_btns_lay.addStretch()

    def _update_save_dir_label(self) -> None:
        save_dir = (
            self.app_state.get("saveDir")
            or data_store.get_setting("defaultSaveDir")
            or data_store.get_desktop_path()
        )
        if save_dir:
            self.app_state["saveDir"] = save_dir
        display = save_dir or "Не выбрана"
        self.lbl_save_dir.setText(f"📁 {display}")
        self.lbl_save_dir.setToolTip(display)

    def _on_choose_save_dir(self) -> None:
        base = self.app_state.get("saveDir") or data_store.get_desktop_path()
        chosen = QFileDialog.getExistingDirectory(
            self, "Выберите папку для сохранения файлов", base
        )
        if chosen:
            self.app_state["saveDir"] = os.path.normpath(chosen)
            data_store.set_setting("defaultSaveDir", self.app_state["saveDir"])
            self._update_save_dir_label()

    def _on_open_save_mode_settings(self) -> None:
        from ui.pages.save_mode_settings_dialog import open_save_mode_settings_dialog
        open_save_mode_settings_dialog(self.window())

    def _get_checked_dept_idx(self) -> int:
        for i, btn in enumerate(self._dept_buttons):
            if btn.isChecked():
                return i
        return 0

    def _on_dept_clicked(self, index: int) -> None:
        if index < 0 or index >= len(self._dept_tab_defs):
            return
        for i, btn in enumerate(self._dept_buttons):
            btn.setChecked(i == index)
        tab_def = self._dept_tab_defs[index]
        scopes = tab_def.get("scopes", [])
        self._clear_subdept_tabs()
        prod_map = data_store.get_products_map()
        if len(scopes) > 1:
            for label, group_idx in scopes:
                g = self._dept_groups[group_idx]
                tab_widget = self._make_dept_tab_widget(g, prod_map)
                self._subdept_tabs_widget.addTab(tab_widget, label)
            self._content_stack.setCurrentIndex(2)
        else:
            if scopes:
                idx, g = scopes[0][1], self._dept_groups[scopes[0][1]]
                self._show_dept_table(idx, g)
                self._content_stack.setCurrentIndex(1)

    def _make_dept_tab_widget(self, group: dict, prod_map: dict) -> QWidget:
        """Создаёт виджет вкладки: подсказка, таблица, кнопка сохранения."""
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(6, 6, 6, 6)
        lay.setSpacing(6)
        templates = data_store.get_ref("templates") or []
        tmpl_name = excel_generator.get_template_name_for_dept(group["key"], templates)
        lbl = QLabel(f"{group['name']}  |  Маршрутов: {len(group['routes'])}  |  Шаблон: {tmpl_name}")
        lbl.setObjectName("hintLabel")
        lay.addWidget(lbl)
        table = QTableWidget()
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.verticalHeader().setVisible(False)
        table.setAlternatingRowColors(True)
        table.setFont(QFont("", self._table_font_size))
        table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        table.installEventFilter(self)
        self._dept_tables.append(table)
        lay.addWidget(table, 1)
        self._fill_dept_table(table, group, prod_map)
        btn = QPushButton(f"Сохранить файл для «{group['name']}»")
        btn.setObjectName("btnSecondary")
        btn.clicked.connect(lambda: self._on_save_single(group))
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_row.addWidget(btn)
        lay.addLayout(btn_row)
        return w

    def _show_dept_table(self, group_idx: int, group: dict) -> None:
        """Создаёт новую таблицу и заполняет данными выбранного отдела/подотдела."""
        self._current_group_idx = group_idx
        prod_map = data_store.get_products_map()
        templates = data_store.get_ref("templates") or []
        tmpl_name = excel_generator.get_template_name_for_dept(group["key"], templates)
        self._dept_hint_label.setText(f"{group['name']}  |  Маршрутов: {len(group['routes'])}  |  Шаблон: {tmpl_name}")

        # Удаляем старую таблицу и создаём новую — устраняет кэширование Qt
        if self._dept_table is not None:
            self._dept_table.removeEventFilter(self)
            self._tbl_lay.removeWidget(self._dept_table)
            self._dept_table.deleteLater()
            self._dept_table = None

        table = QTableWidget()
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.verticalHeader().setVisible(False)
        table.setAlternatingRowColors(True)
        table.setFont(QFont("", self._table_font_size))
        table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        table.installEventFilter(self)
        self._tbl_lay.insertWidget(1, table, 1)
        self._dept_table = table
        self._fill_dept_table(table, group, prod_map)
        self._dept_save_btn.setText(f"Сохранить файл для «{group['name']}»")

    def _on_save_current_dept(self) -> None:
        """Сохраняет файл для текущего выбранного отдела/подотдела."""
        if 0 <= self._current_group_idx < len(self._dept_groups):
            self._on_save_single(self._dept_groups[self._current_group_idx])

    # ─────────────────────────── Рендер ───────────────────────────────────

    def _get_routes_date_str(self) -> str:
        """Дата из app_state (задаётся при добавлении файлов) или завтра."""
        s = self.app_state.get("routesDate")
        if s:
            try:
                parts = s.split(".")
                if len(parts) == 3:
                    return s
            except (ValueError, TypeError):
                pass
        tomorrow = date.today() + timedelta(days=1)
        return f"{tomorrow.day:02d}.{tomorrow.month:02d}.{tomorrow.year}"

    def refresh(self):
        self._refreshing = True
        try:
            self._refresh_impl()
        finally:
            self._refreshing = False

    def _refresh_impl(self):
        file_type = self.app_state.get("fileType", "main")
        self.lbl_title.setText(
            f"Маршруты по отделам — "
            f"{'Основной' if file_type == 'main' else 'Увеличение (Довоз)'}"
        )

        # ── Проверка: есть ли продукты без отдела (4A: баннер + disabled) ────
        unassigned = self._check_unassigned_products()
        if unassigned:
            names_str = ", ".join(unassigned[:8])
            if len(unassigned) > 8:
                names_str += f" и ещё {len(unassigned) - 8}"
            self.lbl_banner.setText(
                f"⚠ {len(unassigned)} продукт(ов) без отдела: {names_str}. "
                "Сначала привяжите все продукты к отделам."
            )
            self.banner_unassigned.setVisible(True)
            self.btn_generate_all.setEnabled(False)
            self.btn_generate_all.setToolTip("Сначала привяжите все продукты к отделам в меню «Справочники» → «Отделы и продукты»")
        else:
            self.banner_unassigned.setVisible(False)
            self.btn_generate_all.setEnabled(True)
            self.btn_generate_all.setToolTip("")

        self._dept_groups = self._build_dept_groups()
        self._dept_tab_defs = self._build_dept_tab_defs()

        self._clear_dept_buttons()
        self._clear_subdept_tabs()

        if not self._dept_groups:
            self._update_save_dir_label()
            self._content_stack.setCurrentIndex(0)
            return

        self._content_stack.setCurrentIndex(1)
        self._populate_dept_buttons()
        self._update_save_dir_label()
        if self._dept_tab_defs:
            self._on_dept_clicked(0)

    def _on_empty_open_last(self) -> None:
        self.go_open_last.emit()

    def _on_empty_process_files(self) -> None:
        self.go_process_files.emit()

    def _on_banner_open_departments(self):
        dept_mod.open_modal(self.window(), self.app_state)
        self.refresh()

    def _on_banner_open_products(self):
        from ui.pages.products_page import open_modal as open_products
        open_products(self.window(), self.app_state)
        self.refresh()

    def _open_depts_and_retry(self):
        dept_mod.open_modal(self, self.app_state)
        self.refresh()

    def eventFilter(self, obj, event):
        """Ctrl + колёсико мыши — масштаб текста в таблицах предпросмотра."""
        if event.type() != QEvent.Type.Wheel or not (QApplication.keyboardModifiers() & Qt.KeyboardModifier.ControlModifier):
            return super().eventFilter(obj, event)
        tables = list(self._dept_tables) if self._dept_tables else ([self._dept_table] if self._dept_table else [])
        if obj not in tables:
            return super().eventFilter(obj, event)
        delta = event.angleDelta().y()
        step = 1 if delta > 0 else -1
        self._table_font_size = max(8, min(24, self._table_font_size + step))
        prod_map = data_store.get_products_map()
        for tbl in tables:
            tbl.setFont(QFont("", self._table_font_size))
        if self._dept_tables:
            dept_idx = self._get_checked_dept_idx()
            if dept_idx < len(self._dept_tab_defs):
                scopes = self._dept_tab_defs[dept_idx].get("scopes", [])
                for i, tbl in enumerate(self._dept_tables):
                    if i < len(scopes):
                        g = self._dept_groups[scopes[i][1]]
                        self._fill_dept_table(tbl, g, prod_map)
        elif self._dept_table and 0 <= self._current_group_idx < len(self._dept_groups):
            self._fill_dept_table(self._dept_table, self._dept_groups[self._current_group_idx], prod_map)
        return True

    def _fill_dept_table(self, table: QTableWidget, group: dict, prod_map: dict):
        """
        Заполняет таблицу данными отдела в формате шаблона (как в сохраняемом файле).
        """
        templates = data_store.get_ref("templates") or []
        sort_asc = self.app_state.get("sortAsc", True)
        headers, header_row2, rows, bold_cells = excel_generator.get_dept_preview_data(
            group, prod_map, templates, sort_asc=sort_asc
        )
        n_cols = len(headers)
        table.setColumnCount(n_cols)
        table.setHorizontalHeaderLabels(headers)
        hdr = table.horizontalHeader()
        hdr.setMinimumSectionSize(60)
        for i in range(n_cols):
            if i == n_cols - 1:
                hdr.setSectionResizeMode(i, QHeaderView.ResizeMode.Stretch)
            else:
                hdr.setSectionResizeMode(i, QHeaderView.ResizeMode.Interactive)
        hdr.resizeSection(0, 90)
        if n_cols > 1:
            hdr.resizeSection(1, 260)
        if n_cols > 2:
            hdr.resizeSection(2, 120)
        if n_cols > 3:
            for j in range(3, n_cols - 1):
                hdr.resizeSection(j, 100)
        total_rows = len(rows)
        if header_row2 is not None:
            total_rows += 1
        table.setUpdatesEnabled(False)
        table.setRowCount(total_rows)

        bold_font = QFont()
        bold_font.setBold(True)
        bold_font.setPointSize(self._table_font_size)
        gray_bg = QBrush(QColor("#f8fafc"))
        FLAG_NO_EDIT = ~Qt.ItemFlag.ItemIsEditable

        def _make_item(text: str, bold: bool = False, bg: QBrush | None = None, align_right: bool = False) -> QTableWidgetItem:
            item = QTableWidgetItem(str(text) if text is not None else "")
            item.setFlags(item.flags() & FLAG_NO_EDIT)
            if bold:
                item.setFont(bold_font)
            if bg is not None:
                item.setBackground(bg)
            if align_right:
                item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            return item

        row_idx = 0
        if header_row2 is not None:
            for ci, val in enumerate(header_row2):
                if ci < n_cols:
                    table.setItem(row_idx, ci, _make_item(val))
            row_idx += 1
        def _looks_numeric(s) -> bool:
            if not s:
                return False
            t = str(s).strip().replace(",", ".")
            try:
                float(t)
                return True
            except ValueError:
                return False

        # bold_cells использует индексы строк из массива rows (0-based), а не индексы строк таблицы.
        # При наличии header_row2 первая строка данных в таблице — row_idx=1, в bold_cells — 0.
        data_row_offset = 1 if header_row2 is not None else 0
        for row_vals in rows:
            data_row_idx = row_idx - data_row_offset
            for ci, val in enumerate(row_vals):
                if ci < n_cols:
                    is_bold = (data_row_idx, ci) in bold_cells
                    table.setItem(row_idx, ci, _make_item(val, bold=is_bold, align_right=_looks_numeric(val)))
            row_idx += 1

        table.setUpdatesEnabled(True)
        # При многих столбцах (шаблон 2, productsWide) не растягивать колонки до контента —
        # иначе таблица становится слишком широкой и возможны сбои отображения.
        if n_cols <= 12:
            table.resizeColumnsToContents()
        row_h = 28
        hdr_h = table.horizontalHeader().height()
        # При многих строках ограничиваем минимальную высоту, чтобы таблица прокручивалась внутри,
        # а не раздувала контент страницы (сбой отображения при шаблоне 2).
        max_visible_rows = 25
        min_h = hdr_h + min(total_rows, max_visible_rows) * row_h + 4
        table.setMinimumHeight(min_h)
        table.setMinimumWidth(200)

    # ─────────────────────────── Генерация ────────────────────────────────

    def _on_save_single(self, group: dict):
        date_str = self._get_routes_date_str()
        self.app_state["routesDate"] = date_str
        file_type = self.app_state.get("fileType", "main")
        base_dir = self.app_state.get("saveDir") or data_store.get_desktop_path()
        data_store.save_last_routes(
            file_type,
            self.app_state.get("routes", []),
            self.app_state.get("uniqueProducts", []),
            self.app_state.get("filteredRoutes", []),
            route_category=self.app_state.get("routeCategory"),
            save_dir=base_dir,
        )
        parent = group.get("parent_dept_name") if group.get("is_subdept") else None
        route_category = None
        if file_type == "main":
            route_category = self.app_state.get("routeCategory")
            if not route_category and group.get("routes"):
                route_category = group["routes"][0].get("routeCategory", "ШК")
        save_path = excel_generator.get_dept_routes_path(
            base_dir, file_type, group["name"], date_str,
            parent_dept_name=parent, route_category=route_category,
        )
        os.makedirs(os.path.dirname(save_path), exist_ok=True)

        try:
            prod_map  = data_store.get_products_map()
            templates = data_store.get_ref("templates") or []
            sort_asc  = self.app_state.get("sortAsc", True)
            replacements = self.app_state.get("productReplacements") or []
            all_routes = [r for r in self.app_state.get("filteredRoutes", []) if not r.get("excluded")]
            replaced_routes = excel_generator.apply_replacements(all_routes, replacements, sort_asc)
            dept_groups = data_store.build_dept_groups_from_routes(replaced_routes)
            target_group = next(
                (g for g in dept_groups if g.get("key") == group.get("key") and g.get("is_subdept") == group.get("is_subdept")),
                group
            )
            excel_generator.generate_single_dept_file(
                target_group, file_type, save_path, prod_map, templates, sort_asc,
                replacements=replacements
            )
            day_dir = excel_generator.get_routes_day_folder(base_dir, date_str)
            products_ref = data_store.get_ref("products") or []
            excel_generator.generate_pcs_compare_report(
                day_dir,
                main_routes=replaced_routes if file_type == "main" else [],
                increase_routes=replaced_routes if file_type == "increase" else [],
                products_ref=products_ref,
                date_str=date_str,
            )
            QMessageBox.information(self, "Готово", f"Файл создан или обновлён:\n{save_path}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при создании файла:\n{e}")

    def _on_generate_all(self):
        if not self._dept_groups:
            QMessageBox.warning(self, "Нет данных",
                                "Нет отделов с маршрутами для генерации.")
            return

        chosen_dir = self.app_state.get("saveDir") or data_store.get_desktop_path()

        prod_map  = data_store.get_products_map()
        templates = data_store.get_ref("templates") or []
        file_type = self.app_state.get("fileType", "main")
        data_store.save_last_routes(
            file_type,
            self.app_state.get("routes", []),
            self.app_state.get("uniqueProducts", []),
            self.app_state.get("filteredRoutes", []),
            route_category=self.app_state.get("routeCategory"),
            save_dir=chosen_dir,
        )
        sort_asc  = self.app_state.get("sortAsc", True)
        date_str = self._get_routes_date_str()
        self.app_state["routesDate"] = date_str
        if file_type == "main":
            for cat in ("ШК", "СД"):
                type_dir = excel_generator.get_routes_type_folder(chosen_dir, file_type, date_str, route_category=cat)
                os.makedirs(type_dir, exist_ok=True)
        else:
            type_dir = excel_generator.get_routes_type_folder(chosen_dir, file_type, date_str)
            os.makedirs(type_dir, exist_ok=True)
        general_path = excel_generator.get_general_routes_path(chosen_dir, file_type, date_str) if file_type != "main" else None
        all_routes = [r for r in self.app_state.get("filteredRoutes", []) if not r.get("excluded")]

        self.btn_generate_all.setEnabled(False)
        self.progress.setVisible(True)
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)

        replacements = self.app_state.get("productReplacements") or []
        self._gen_thread = QThread(self)
        self._gen_worker = DeptGenerateWorker(
            self._dept_groups, file_type, chosen_dir, prod_map, templates, sort_asc,
            all_routes=all_routes, general_path=general_path, date_str=date_str,
            replacements=replacements,
        )
        self._gen_worker.moveToThread(self._gen_thread)
        self._gen_thread.started.connect(self._gen_worker.run)
        self._gen_worker.finished.connect(self._on_gen_done)
        self._gen_worker.error.connect(self._on_gen_error)
        self._gen_worker.finished.connect(self._gen_thread.quit)
        self._gen_worker.error.connect(self._gen_thread.quit)
        self._gen_thread.start()

    def _on_gen_done(self, created: list):
        QApplication.restoreOverrideCursor()
        self.progress.setVisible(False)
        self.btn_generate_all.setEnabled(True)
        self.app_state["generalFileCreated"] = True
        self.app_state["deptFilesCreated"] = True
        set_status = self.app_state.get("set_status")
        if callable(set_status):
            set_status(f"Создано {len(created)} файлов")
        upd = self.app_state.get("_update_tabs")
        if callable(upd):
            upd()
        QMessageBox.information(
            self, "Готово",
            f"Файлы созданы или обновлены: {len(created)}\n\n" +
            "\n".join(os.path.basename(p) for p in created[:10]) +
            ("\n..." if len(created) > 10 else "")
        )

    def _on_gen_error(self, msg: str):
        QApplication.restoreOverrideCursor()
        self.progress.setVisible(False)
        self.btn_generate_all.setEnabled(True)
        QMessageBox.critical(self, "Ошибка", f"Ошибка при создании файлов:\n{msg}")

