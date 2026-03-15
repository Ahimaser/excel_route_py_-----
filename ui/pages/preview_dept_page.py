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
    QProgressBar, QTabWidget, QApplication, QScrollArea,
    QCheckBox, QSizePolicy,
)
from PyQt6.QtCore import Qt, pyqtSignal, QThread, QObject, QEvent, QTimer
from PyQt6.QtGui import QFont, QBrush, QColor, QWheelEvent

from core import data_store, excel_generator
from ui.pages import departments_page as dept_mod
from ui.excel_safe_open import open_excel_file_safely


# ─────────────────────────── Worker ───────────────────────────────────────

class DeptGenerateWorker(QObject):
    finished = pyqtSignal(list)
    error    = pyqtSignal(str)

    def __init__(self, dept_groups: list, file_type: str, save_dir: str,
                 prod_map: dict, templates: list, sort_asc: bool = True,
                 all_routes: list | None = None, general_path: str | None = None,
                 date_str: str | None = None, replacements: list | None = None,
                 split_by_product: bool = False, product_groups: dict | None = None):
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
        self.split_by_product = split_by_product
        self.product_groups = product_groups or {}

    def run(self):
        try:
            replaced_routes = excel_generator.apply_replacements(
                self.all_routes, self.replacements, self.sort_asc
            )
            dept_groups = data_store.build_dept_groups_from_routes(replaced_routes)
            if self.general_path:
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
            if self.split_by_product:
                created = excel_generator.generate_dept_files_by_products(
                    dept_groups, self.product_groups, self.file_type,
                    self.save_dir, self.prod_map, self.templates,
                    self.sort_asc, date_str=self.date_str,
                    replacements=self.replacements
                )
            else:
                created = excel_generator.generate_dept_files(
                    dept_groups, self.file_type,
                    self.save_dir, self.prod_map, self.templates,
                    self.sort_asc, date_str=self.date_str,
                    replacements=self.replacements
                )
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
            if self.general_path:
                created = [self.general_path, *created]
            self.finished.emit(created)
        except Exception as e:
            self.error.emit(str(e))


class DeptTemplateLabelsWorker(QObject):
    finished = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(self, routes: list, base_dir: str, file_type: str, products_ref: list,
                 departments_ref: list, date_str: str | None = None):
        super().__init__()
        import copy
        self.routes = copy.deepcopy(routes)
        self.base_dir = base_dir
        self.file_type = file_type
        self.products_ref = products_ref
        self.departments_ref = departments_ref
        self.date_str = date_str

    def run(self):
        try:
            created = excel_generator.generate_simple_labels(
                self.routes,
                self.base_dir,
                self.file_type,
                self.products_ref,
                self.departments_ref,
                date_str=self.date_str,
            )
            self.finished.emit(created)
        except Exception as exc:
            self.error.emit(str(exc))


# ─────────────────────────── Страница ─────────────────────────────────────

class PreviewDeptPage(QWidget):
    """Предпросмотр и генерация файлов по отделам."""

    go_back = pyqtSignal()
    go_home = pyqtSignal()  # Переход на главную (dashboard)
    go_clear_routes = pyqtSignal()

    def __init__(self, app_state: dict):
        super().__init__()
        self.app_state   = app_state
        self._dept_groups: list[dict] = []
        self._table_font_size = 11
        self._dept_tables: list[QTableWidget] = []
        self._col_widths_by_table = {}
        self._dept_tab_defs: list[dict] = []
        self._dept_buttons: list[QPushButton] = []
        self._subdept_buttons: list[QPushButton] = []
        self._subdept_scopes: list[tuple[str, int]] = []
        self._refreshing = False
        self._build_ui()

    def _build_ui(self):
        self._scroll_area = QScrollArea()
        scroll = self._scroll_area
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        content = QWidget()
        main_lay = QVBoxLayout(self)
        main_lay.setContentsMargins(0, 0, 0, 0)
        scroll.setWidget(content)
        content.setObjectName("previewDeptContent")
        content.setMinimumHeight(480)
        inner = QVBoxLayout(content)
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

        self.subdept_frame = QFrame()
        self.subdept_frame.setObjectName("subdeptPillsBar")
        self.subdept_btns_lay = QHBoxLayout(self.subdept_frame)
        self.subdept_btns_lay.setContentsMargins(0, 4, 0, 4)
        self.subdept_btns_lay.setSpacing(6)
        self.subdept_frame.setVisible(False)
        inner.addWidget(self.subdept_frame)

        self.tabs = QTabWidget()
        self.tabs.tabBar().hide()
        self.tabs.currentChanged.connect(self._on_dept_tab_changed)
        inner.addWidget(self.tabs, 1)

        main_lay.addWidget(scroll, 1)

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
        self.chk_split_by_product = QCheckBox("Разделить по продуктам")
        self.chk_split_by_product.setChecked(False)
        self.chk_split_by_product.setToolTip(
            "Создавать отдельный файл на каждый продукт или группу. "
            "Доступно только для отделов с Шаблоном 2 (Компактный). "
            "Включите, чтобы появилась кнопка «Настроить группы»."
        )
        self.chk_split_by_product.stateChanged.connect(self._on_split_by_product_changed)
        bottom_row.addWidget(self.chk_split_by_product)
        self.btn_config_groups = QPushButton("Настроить группы")
        self.btn_config_groups.setObjectName("btnSecondary")
        self.btn_config_groups.setToolTip(
            "Объединить несколько продуктов в один файл. "
            "Сначала включите «Разделить по продуктам»."
        )
        self.btn_config_groups.clicked.connect(self._on_config_product_groups)
        self.btn_config_groups.setEnabled(False)
        bottom_row.addWidget(self.btn_config_groups)
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

    def _clear_subdept_buttons(self) -> None:
        while self.subdept_btns_lay.count():
            item = self.subdept_btns_lay.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
        self._subdept_buttons.clear()
        self._subdept_scopes.clear()
        self.subdept_frame.setVisible(False)

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

    def _on_split_by_product_changed(self, state) -> None:
        self.btn_config_groups.setEnabled(state == Qt.CheckState.Checked)

    def _on_config_product_groups(self) -> None:
        """Открывает диалог настройки групп продуктов для текущего отдела."""
        idx = self.tabs.currentIndex()
        if idx < 0 or idx >= len(self._dept_groups):
            QMessageBox.warning(self, "Нет отдела", "Выберите отдел для настройки групп.")
            return
        group = self._dept_groups[idx]
        dept_key = group.get("key", "")
        dept_name = group.get("name", "")
        products = data_store.get_ref("products") or []
        product_names = [p["name"] for p in products if p.get("deptKey") == dept_key]
        if not product_names:
            QMessageBox.information(
                self, "Нет продуктов",
                f"У отдела «{dept_name}» нет привязанных продуктов."
            )
            return
        from ui.pages.product_groups_dialog import open_product_groups_dialog
        open_product_groups_dialog(self.window(), dept_key, dept_name, product_names)

    def _on_dept_clicked(self, index: int) -> None:
        if index < 0 or index >= len(self._dept_tab_defs):
            return
        for i, btn in enumerate(self._dept_buttons):
            btn.setChecked(i == index)
        tab_def = self._dept_tab_defs[index]
        scopes = tab_def.get("scopes", [])
        self._clear_subdept_buttons()
        if len(scopes) > 1:
            self.subdept_frame.setVisible(True)
            self._subdept_scopes = scopes
            for i, (label, group_idx) in enumerate(scopes):
                btn = QPushButton(label)
                btn.setObjectName("subdeptPill")
                btn.setCheckable(True)
                btn.setChecked(i == 0)
                btn.clicked.connect(lambda c=False, idx=group_idx: self._on_scope_clicked(idx))
                self.subdept_btns_lay.addWidget(btn)
                self._subdept_buttons.append(btn)
            self.subdept_btns_lay.addStretch()
            self.tabs.blockSignals(True)
            self.tabs.setCurrentIndex(scopes[0][1])
            self.tabs.blockSignals(False)
        else:
            self.subdept_frame.setVisible(False)
            if scopes:
                self.tabs.blockSignals(True)
                self.tabs.setCurrentIndex(scopes[0][1])
                self.tabs.blockSignals(False)

    def _on_scope_clicked(self, group_idx: int) -> None:
        for i, (_, idx) in enumerate(self._subdept_scopes):
            if i < len(self._subdept_buttons):
                self._subdept_buttons[i].setChecked(idx == group_idx)
        self.tabs.blockSignals(True)
        self.tabs.setCurrentIndex(group_idx)
        self.tabs.blockSignals(False)

    def _on_dept_tab_changed(self, index: int) -> None:
        """При переключении вкладки обновляет отображение таблицы."""
        if self._refreshing or index < 0:
            return
        if index >= len(self._dept_tables):
            return
        w = self.tabs.widget(index)
        if w:
            w.show()
        tbl = self._dept_tables[index]
        tbl.viewport().update()
        tbl.update()

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

        # Разделение по продуктам — только если есть отделы с Шаблоном 2
        templates = data_store.get_ref("templates") or []
        has_template2 = any(
            excel_generator.dept_has_template_2(g.get("key", ""), templates)
            for g in self._dept_groups
        )
        self.chk_split_by_product.setEnabled(has_template2)
        if not has_template2 and self.chk_split_by_product.isChecked():
            self.chk_split_by_product.setChecked(False)
            self.btn_config_groups.setEnabled(False)

        self.tabs.clear()
        self._dept_tables.clear()
        self._clear_dept_buttons()
        self._clear_subdept_buttons()

        if not self._dept_groups:
            self._update_save_dir_label()
            empty = QLabel(
                "Нет данных для отображения.\n\n"
                "Убедитесь, что:\n"
                "  • Загружены XLS файлы\n"
                "  • Созданы отделы с привязанными продуктами"
            )
            empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
            empty.setObjectName("stepLabel")
            self.tabs.addTab(empty, "Нет данных")
            return

        prod_map = data_store.get_products_map()

        for group in self._dept_groups:
            tab = self._make_dept_tab(group, prod_map)
            self.tabs.addTab(tab, group["name"])

        self._populate_dept_buttons()
        self._update_save_dir_label()
        if self._dept_tab_defs:
            self._on_dept_clicked(0)

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

    def _make_dept_tab(self, group: dict, prod_map: dict) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(6, 6, 6, 6)
        lay.setSpacing(6)

        templates = data_store.get_ref("templates") or []
        tmpl_name = excel_generator.get_template_name_for_dept(group["key"], templates)
        lbl = QLabel(f"Маршрутов: {len(group['routes'])}  |  Шаблон: {tmpl_name}")
        lbl.setObjectName("hintLabel")
        lay.addWidget(lbl)

        sort_asc = self.app_state.get("sortAsc", True)
        headers, header_row2, rows, _ = excel_generator.get_dept_preview_data(
            group, prod_map, templates, sort_asc=sort_asc
        )
        n_cols = len(headers)
        table = QTableWidget()
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
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.verticalHeader().setVisible(False)
        table.setAlternatingRowColors(True)
        table.setFont(QFont("", self._table_font_size))
        table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        table.installEventFilter(self)
        self._dept_tables.append(table)
        lay.addWidget(table, 1)

        self._fill_dept_table(table, group, prod_map)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_save = QPushButton(f"Сохранить файл для «{group['name']}»")
        btn_save.setObjectName("btnSecondary")
        btn_save.clicked.connect(lambda _, g=group: self._on_save_single(g))
        btn_row.addWidget(btn_save)
        lay.addLayout(btn_row)

        return w

    def eventFilter(self, obj, event):
        """Ctrl + колёсико мыши — масштаб текста в таблицах предпросмотра."""
        if event.type() == QEvent.Type.Wheel and obj in self._dept_tables:
            if QApplication.keyboardModifiers() & Qt.KeyboardModifier.ControlModifier:
                delta = event.angleDelta().y()
                step = 1 if delta > 0 else -1
                self._table_font_size = max(8, min(24, self._table_font_size + step))
                prod_map = data_store.get_products_map()
                for i, tbl in enumerate(self._dept_tables):
                    tbl.setFont(QFont("", self._table_font_size))
                    if i < len(self._dept_groups):
                        self._fill_dept_table(tbl, self._dept_groups[i], prod_map)
                return True
        return super().eventFilter(obj, event)

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
        if table.columnCount() != n_cols:
            table.setColumnCount(n_cols)
            table.setHorizontalHeaderLabels(headers)
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

        for row_vals in rows:
            for ci, val in enumerate(row_vals):
                if ci < n_cols:
                    is_bold = (row_idx, ci) in bold_cells
                    table.setItem(row_idx, ci, _make_item(val, bold=is_bold, align_right=_looks_numeric(val)))
            row_idx += 1

        table.setUpdatesEnabled(True)
        table.resizeColumnsToContents()
        row_h = 28
        hdr_h = table.horizontalHeader().height()
        table.setMinimumHeight(hdr_h + total_rows * row_h + 4)
        # Таблица растягивается на всю ширину контейнера (последний столбец — Stretch)
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
        save_path = excel_generator.get_dept_routes_path(
            base_dir, file_type, group["name"], date_str, parent_dept_name=parent
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
        type_dir = excel_generator.get_routes_type_folder(chosen_dir, file_type, date_str)
        os.makedirs(type_dir, exist_ok=True)
        general_path = excel_generator.get_general_routes_path(chosen_dir, file_type, date_str)
        all_routes = [r for r in self.app_state.get("filteredRoutes", []) if not r.get("excluded")]

        self.btn_generate_all.setEnabled(False)
        self.progress.setVisible(True)
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)

        replacements = self.app_state.get("productReplacements") or []
        split_by_product = self.chk_split_by_product.isChecked()
        product_groups = data_store.get_setting("productFileGroups") or {} if split_by_product else {}
        self._gen_thread = QThread(self)
        self._gen_worker = DeptGenerateWorker(
            self._dept_groups, file_type, chosen_dir, prod_map, templates, sort_asc,
            all_routes=all_routes, general_path=general_path, date_str=date_str,
            replacements=replacements,
            split_by_product=split_by_product,
            product_groups=product_groups,
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

    def _on_labels_from_templates(self):
        routes = self.app_state.get("filteredRoutes", [])
        active = [r for r in routes if not r.get("excluded")]
        if not active:
            QMessageBox.warning(self, "Нет данных", "Нет маршрутов для этикеток.")
            return
        products_ref = data_store.get_ref("products") or []
        departments_ref = data_store.get_ref("departments") or []
        base_dir = self.app_state.get("saveDir") or data_store.get_desktop_path()
        file_type = self.app_state.get("fileType", "main")
        date_str = self._get_routes_date_str()
        type_dir = excel_generator.get_routes_type_folder(base_dir, file_type, date_str)
        self._labels_out_dir = os.path.join(type_dir, f"Этикетки на {date_str}")
        self.btn_labels.setEnabled(False)
        self.progress.setVisible(True)
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        self._labels_thread = QThread(self)
        self._labels_worker = DeptTemplateLabelsWorker(
            routes, base_dir, file_type, products_ref, departments_ref, date_str=date_str
        )
        self._labels_worker.moveToThread(self._labels_thread)
        self._labels_thread.started.connect(self._labels_worker.run)
        self._labels_worker.finished.connect(self._on_template_labels_done)
        self._labels_worker.error.connect(self._on_template_labels_error)
        self._labels_worker.finished.connect(self._labels_thread.quit)
        self._labels_worker.error.connect(self._labels_thread.quit)
        self._labels_thread.finished.connect(self._labels_worker.deleteLater)
        self._labels_thread.finished.connect(self._labels_thread.deleteLater)
        self._labels_thread.start()

    def _on_template_labels_done(self, created: list):
        QApplication.restoreOverrideCursor()
        self.progress.setVisible(False)
        self.btn_labels.setEnabled(True)
        if created:
            set_status = self.app_state.get("set_status")
            if callable(set_status):
                set_status(f"Создано этикеток: {len(created)}")
            QMessageBox.information(self, "Готово", f"Создано файлов: {len(created)}\n\n{self._labels_out_dir}")
            reply = QMessageBox.question(
                self,
                "Открыть безопасно",
                "Открыть первый созданный файл через безопасную локальную копию?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes,
            )
            if reply == QMessageBox.StandardButton.Yes:
                try:
                    open_excel_file_safely(created[0])
                except Exception as exc:
                    QMessageBox.warning(self, "Не удалось открыть файл", str(exc))
        else:
            QMessageBox.information(self, "Нет файлов", "Нет этикеток для создания.")

    def _on_template_labels_error(self, msg: str):
        QApplication.restoreOverrideCursor()
        self.progress.setVisible(False)
        self.btn_labels.setEnabled(True)
        QMessageBox.critical(self, "Ошибка", msg)

