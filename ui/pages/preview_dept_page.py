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
)
from PyQt6.QtCore import Qt, pyqtSignal, QThread, QObject, QEvent
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
                 date_str: str | None = None):
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

    def run(self):
        try:
            if self.general_path:
                os.makedirs(os.path.dirname(self.general_path), exist_ok=True)
                excel_generator.generate_general_routes(
                    self.all_routes,
                    self.file_type,
                    self.general_path,
                    self.prod_map,
                    self.sort_asc,
                    date_str=self.date_str,
                )
            created = excel_generator.generate_dept_files(
                self.dept_groups, self.file_type,
                self.save_dir, self.prod_map, self.templates,
                self.sort_asc, date_str=self.date_str
            )
            if self.general_path:
                created = [self.general_path, *created]
            self.finished.emit(created)
        except Exception as e:
            self.error.emit(str(e))


class DeptTemplateLabelsWorker(QObject):
    finished = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(self, routes: list, out_dir: str, file_type: str, products_ref: list, departments_ref: list):
        super().__init__()
        import copy
        self.routes = copy.deepcopy(routes)
        self.out_dir = out_dir
        self.file_type = file_type
        self.products_ref = products_ref
        self.departments_ref = departments_ref

    def run(self):
        try:
            created = excel_generator.generate_labels_from_templates(
                self.routes, self.out_dir, self.file_type, self.products_ref, self.departments_ref
            )
            self.finished.emit(created)
        except Exception as exc:
            self.error.emit(str(exc))


# ─────────────────────────── Страница ─────────────────────────────────────

_DISPLAY_ADDR_ONLY = "addr"
_DISPLAY_FULL = "full"


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
        self._display_mode = _DISPLAY_ADDR_ONLY  # по умолчанию только № маршрута и адрес
        self._dept_tab_defs: list[dict] = []
        self._dept_buttons: list[QPushButton] = []
        self._subdept_buttons: list[QPushButton] = []
        self._subdept_scopes: list[tuple[str, int]] = []
        self._build_ui()

    def _build_ui(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        content = QWidget()
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.addWidget(scroll)
        scroll.setWidget(content)
        content.setObjectName("previewDeptContent")
        content.setMinimumHeight(480)
        inner = QVBoxLayout(content)
        inner.setContentsMargins(20, 16, 20, 16)
        inner.setSpacing(12)

        h_row = QHBoxLayout()
        btn_back = QPushButton("← Назад")
        btn_back.setObjectName("btnBack")
        btn_back.clicked.connect(self.go_back.emit)
        h_row.addWidget(btn_back)
        btn_home = QPushButton("На главную")
        btn_home.setObjectName("btnSecondary")
        btn_home.clicked.connect(self.go_home.emit)
        h_row.addWidget(btn_home)

        self.lbl_title = QLabel("Маршруты по отделам")
        self.lbl_title.setObjectName("sectionTitle")
        h_row.addWidget(self.lbl_title)
        h_row.addStretch()

        h_row.addWidget(QLabel("Показать:"))
        self.combo_display = QComboBox()
        self.combo_display.setMinimumWidth(140)
        self.combo_display.addItem("Только № и адрес", _DISPLAY_ADDR_ONLY)
        self.combo_display.addItem("Полностью", _DISPLAY_FULL)
        self.combo_display.setCurrentIndex(0)
        self.combo_display.currentIndexChanged.connect(self._on_display_changed)
        h_row.addWidget(self.combo_display)

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
        inner.addWidget(self.tabs)

        bottom_row = QHBoxLayout()
        bottom_row.addStretch()
        self.btn_generate_all = QPushButton("Создать файлы для всех отделов")
        self.btn_generate_all.setMinimumWidth(220)
        self.btn_generate_all.setObjectName("btnPrimary")
        self.btn_generate_all.clicked.connect(self._on_generate_all)
        bottom_row.addWidget(self.btn_generate_all)

        inner.addLayout(bottom_row)

        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.setVisible(False)
        inner.addWidget(self.progress)

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
                        "products":       dept_prods,
                    })
            return result

        for dept in depts:
            # Подотделы
            for sub in dept.get("subdepts", []):
                sub_routes = _collect_routes(sub["key"])
                if sub_routes:
                    groups.append({
                        "key":        sub["key"],
                        "name":       sub["name"],
                        "is_subdept": True,
                        "routes":     sub_routes,
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
            self.tabs.setCurrentIndex(scopes[0][1])
        else:
            self.subdept_frame.setVisible(False)
            if scopes:
                self.tabs.setCurrentIndex(scopes[0][1])

    def _on_scope_clicked(self, group_idx: int) -> None:
        for i, (_, idx) in enumerate(self._subdept_scopes):
            if i < len(self._subdept_buttons):
                self._subdept_buttons[i].setChecked(idx == group_idx)
        self.tabs.setCurrentIndex(group_idx)

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
            self.btn_generate_all.setToolTip("Сначала привяжите все продукты к отделам")
        else:
            self.banner_unassigned.setVisible(False)
            self.btn_generate_all.setEnabled(True)
            self.btn_generate_all.setToolTip("")

        self._dept_groups = self._build_dept_groups()
        self._dept_tab_defs = self._build_dept_tab_defs()

        self.tabs.clear()
        self._dept_tables.clear()
        self._clear_dept_buttons()
        self._clear_subdept_buttons()

        if not self._dept_groups:
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
        if self._dept_tab_defs:
            self._on_dept_clicked(0)

    def _show_blocked(self, unassigned: list):
        """Показывает заглушку, пока есть непривязанные продукты."""
        self._dept_groups = []
        self._dept_tab_defs = []
        self._clear_dept_buttons()
        self._clear_subdept_buttons()
        self.tabs.clear()
        names_str = ", ".join(unassigned[:5])
        if len(unassigned) > 5:
            names_str += f" и ещё {len(unassigned)-5}"
        blocked = QWidget()
        blay = QVBoxLayout(blocked)
        blay.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl = QLabel(
            f"⚠ Страница недоступна\n\n"
            f"Продукты без отдела: {names_str}\n\n"
            f"Откройте «Справочники → Отделы и продукты» и привяжите все продукты к отделам."
        )
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl.setObjectName("warningLabel")
        lbl.setWordWrap(True)
        blay.addWidget(lbl)
        btn_row = QHBoxLayout()
        btn_row.setSpacing(12)
        btn_open = QPushButton("Открыть Отделы и продукты")
        btn_open.setMinimumWidth(200)
        btn_open.setObjectName("btnPrimary")
        btn_open.clicked.connect(lambda: self._open_depts_and_retry())
        btn_retry = QPushButton("Проверить снова")
        btn_retry.setObjectName("btnSecondary")
        btn_retry.clicked.connect(self.refresh)
        btn_row.addWidget(btn_open, alignment=Qt.AlignmentFlag.AlignCenter)
        btn_row.addWidget(btn_retry, alignment=Qt.AlignmentFlag.AlignCenter)
        blay.addLayout(btn_row)
        self.tabs.addTab(blocked, "Заблокировано")

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

        lbl = QLabel(f"Маршрутов: {len(group['routes'])}")
        lbl.setObjectName("hintLabel")
        lay.addWidget(lbl)

        has_dirty = any(
            prod_map.get(p.get("name", ""), {}).get("showInDirty")
            and data_store.is_subdept_chistchenka(prod_map.get(p.get("name", ""), {}).get("deptKey"))
            for r in group["routes"] for p in r.get("products", [])
        )
        n_cols = 6 if has_dirty else 5
        headers = ["№ маршрута", "Адрес / Продукт", "Ед. изм.", "Кол-во", "Шт"]
        if has_dirty:
            headers.append("Грязные")
        table = QTableWidget()
        table.setColumnCount(n_cols)
        table.setHorizontalHeaderLabels(headers)
        hdr = table.horizontalHeader()
        hdr.setMinimumSectionSize(90)
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        if has_dirty:
            hdr.setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        hdr.resizeSection(1, 260)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.verticalHeader().setVisible(False)
        table.setAlternatingRowColors(True)
        table.setFont(QFont("", self._table_font_size))
        table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        table.installEventFilter(self)
        self._dept_tables.append(table)
        lay.addWidget(table)

        self._fill_dept_table(table, group["routes"], prod_map)

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
                        self._fill_dept_table(tbl, self._dept_groups[i]["routes"], prod_map)
                return True
        return super().eventFilter(obj, event)

    def _on_display_changed(self):
        self._display_mode = self.combo_display.currentData() or _DISPLAY_ADDR_ONLY
        prod_map = data_store.get_products_map()
        for i, tbl in enumerate(self._dept_tables):
            if i < len(self._dept_groups):
                self._fill_dept_table(tbl, self._dept_groups[i]["routes"], prod_map)

    def _fill_dept_table(self, table: QTableWidget, routes: list, prod_map: dict):
        """
        Заполняет таблицу данными отдела (batch-рендер).
        При режиме «Только № и адрес» — только строки маршрутов; иначе — маршрут + продукты.
        """
        sort_asc = self.app_state.get("sortAsc", True)
        from core.excel_generator import _sort_routes
        routes_sorted = _sort_routes(routes, sort_asc)
        only_addr = self._display_mode == _DISPLAY_ADDR_ONLY

        if only_addr:
            total_rows = len(routes_sorted)
        else:
            total_rows = sum(1 + len(r.get("products", [])) for r in routes_sorted)

        table.setUpdatesEnabled(False)
        table.setRowCount(total_rows)

        bold_font = QFont()
        bold_font.setBold(True)
        bold_font.setPointSize(self._table_font_size)
        gray_bg   = QBrush(QColor("#f8fafc"))
        FLAG_NO_EDIT = ~Qt.ItemFlag.ItemIsEditable

        def _make_item(
            text: str,
            bold: bool = False,
            bg: QBrush | None = None,
            align_right: bool = False,
        ) -> QTableWidgetItem:
            item = QTableWidgetItem(text)
            item.setFlags(item.flags() & FLAG_NO_EDIT)
            if bold:
                item.setFont(bold_font)
            if bg is not None:
                item.setBackground(bg)
            if align_right:
                item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            return item

        current_row = 0
        for route in routes_sorted:
            products  = route.get("products", [])
            route_num = str(route.get("routeNum", ""))
            address   = route.get("address", "")

            has_dirty_col = table.columnCount() >= 6
            # Строка маршрута: номер | адрес | пусто | пусто | пусто | [пусто]
            table.setItem(current_row, 0, _make_item(route_num, bold=True))
            table.setItem(current_row, 1, _make_item(address,   bold=True, bg=gray_bg))
            table.setItem(current_row, 2, _make_item(""))
            table.setItem(current_row, 3, _make_item(""))
            table.setItem(current_row, 4, _make_item(""))
            if has_dirty_col:
                table.setItem(current_row, 5, _make_item(""))
            current_row += 1

            if not only_addr:
                # Строки продуктов: пусто | название | ед.изм. | кол-во | шт | [грязные]
                for prod in products:
                    table.setItem(current_row, 0, _make_item(""))
                    table.setItem(current_row, 1, _make_item(f"  {prod.get('name', '')}"))
                    table.setItem(current_row, 2, _make_item(prod.get("unit", "")))
                    qty_str = self._fmt_qty(prod, prod_map)
                    pcs_str = self._fmt_pcs(prod, prod_map, route)
                    table.setItem(current_row, 3, _make_item(qty_str, align_right=True))
                    table.setItem(current_row, 4, _make_item(pcs_str, align_right=True))
                    if has_dirty_col:
                        dirty_str = self._fmt_dirty_qty(prod, prod_map)
                        table.setItem(current_row, 5, _make_item(dirty_str, align_right=True))
                    current_row += 1

        table.setUpdatesEnabled(True)
        table.resizeColumnsToContents()
        row_h = 28
        hdr_h = table.horizontalHeader().height()
        table.setMinimumHeight(hdr_h + total_rows * row_h + 4)

    def _fmt_dirty_qty(self, prod: dict, prod_map: dict) -> str:
        """Кол-во грязных (×1,25) для продуктов с showInDirty в подотделе Чищенка, иначе «—»."""
        ps = prod_map.get(prod.get("name", ""), {})
        if not ps.get("showInDirty") or not data_store.is_subdept_chistchenka(ps.get("deptKey")):
            return "—"
        qty = prod.get("quantity")
        if qty is None:
            return ""
        try:
            dirty = float(qty) * 1.25
            return str(int(dirty)) if abs(dirty - round(dirty)) < 1e-9 else str(round(dirty, 2))
        except (ValueError, TypeError):
            return ""

    def _fmt_qty(self, prod: dict, prod_map: dict) -> str:
        """Кол-во в единицах измерения (с учётом множителя замены)."""
        qty = prod.get("quantity")
        if qty is None:
            return ""
        ps = prod_map.get(prod.get("name", ""), {})
        mult = float(ps.get("quantityMultiplier", 1.0) or 1.0)
        try:
            display_qty = float(qty) * mult
        except (ValueError, TypeError):
            display_qty = qty
        if isinstance(display_qty, float) and display_qty == int(display_qty):
            return str(int(display_qty))
        return str(display_qty)

    def _fmt_pcs(self, prod: dict, prod_map: dict, route: dict) -> str:
        """
        Значение для колонки «Шт»: рассчитанные шт (по настройкам продукта и категории маршрута)
        или количество для продуктов с ед. изм. «шт». Иначе «—».
        """
        qty = prod.get("quantity")
        if qty is None:
            return "—"
        unit = (prod.get("unit") or "").strip().lower()
        ps = prod_map.get(prod.get("name", ""), {})

        if unit == "шт":
            try:
                v = float(qty)
                return str(int(v)) if v == int(v) else str(v)
            except (ValueError, TypeError):
                return str(qty)

        mult = float(ps.get("quantityMultiplier", 1.0) or 1.0)
        try:
            display_qty = float(qty) * mult
        except (ValueError, TypeError):
            display_qty = qty

        if not ps.get("showPcs") or not (ps.get("pcsPerUnit") or 0) > 0:
            return "—"

        try:
            val = float(display_qty)
        except (TypeError, ValueError):
            return "—"
        mode = excel_generator.get_dept_special_mode(ps.get("deptKey"))
        if mode == "polufabricates":
            pcu = float(ps.get("pcsPerUnit", 1) or 1)
            if pcu <= 0:
                return "0"
            pcs = max(0, int(val // pcu))
            tail = max(0.0, val - pcs * pcu)
            unit = (prod.get("unit") or "").strip()
            if tail > 1e-9 and unit:
                tail_txt = str(int(tail)) if abs(tail - round(tail)) < 1e-9 else f"{tail:.3f}".rstrip("0").rstrip(".")
                return f"{pcs} шт + {tail_txt} {unit}"
            return f"{pcs} шт"
        # Порог: ниже — 0 шт (как в генераторе)
        min_qty = ps.get("minQtyForPcs")
        if min_qty is not None and min_qty > 0:
            threshold = float(min_qty)
        else:
            unit_lower = (prod.get("unit") or "").strip().lower()
            threshold = 0.2 if unit_lower in ("кг", "л", "kg", "l") else 0
        if val < threshold:
            return "0"

        route_cat = route.get("routeCategory") or "ШК"
        pcu = float(ps.get("pcsPerUnit", 1))
        addr = route.get("address", "")
        force_round_up = excel_generator.is_always_round_up_institution(addr)
        if force_round_up:
            pct = excel_generator.get_institution_round_percent(ps.get("deptKey"))
            round_tail = pcu * (pct / 100.0)
            pcs = excel_generator.calc_pcs_tail(val, pcu, round_tail)
        else:
            round_tail = ps.get("roundTailFromСД") if route_cat == "СД" else ps.get("roundTailFromШК")
            if round_tail is not None:
                round_tail = float(round_tail)
                pcs = excel_generator.calc_pcs_tail(val, pcu, round_tail)
            else:
                round_up = (
                    ps.get("roundUpСД") if "roundUpСД" in ps else ps.get("roundUp", True)
                    if route_cat == "СД"
                    else ps.get("roundUpШК") if "roundUpШК" in ps else ps.get("roundUp", True)
                )
                pcs = excel_generator.calc_pcs(val, pcu, bool(round_up))
        return str(pcs)

    # ─────────────────────────── Генерация ────────────────────────────────

    def _on_save_single(self, group: dict):
        date_str = self._get_routes_date_str()
        self.app_state["routesDate"] = date_str
        file_type = self.app_state.get("fileType", "main")
        data_store.save_last_routes(
            file_type,
            self.app_state.get("routes", []),
            self.app_state.get("uniqueProducts", []),
            self.app_state.get("filteredRoutes", []),
            route_category=self.app_state.get("routeCategory"),
        )
        base_dir = self.app_state.get("saveDir") or data_store.get_desktop_path()
        chosen_base = QFileDialog.getExistingDirectory(
            self,
            "Выберите папку для сохранения маршрутов",
            base_dir,
        )
        if not chosen_base:
            return
        self.app_state["saveDir"] = chosen_base
        save_path = excel_generator.get_dept_routes_path(chosen_base, file_type, group["name"], date_str)
        os.makedirs(os.path.dirname(save_path), exist_ok=True)

        try:
            prod_map  = data_store.get_products_map()
            templates = data_store.get_ref("templates") or []
            sort_asc  = self.app_state.get("sortAsc", True)
            excel_generator.generate_single_dept_file(
                group, file_type, save_path, prod_map, templates, sort_asc
            )
            day_dir = excel_generator.get_routes_day_folder(chosen_base, date_str)
            report_path = os.path.join(day_dir, f"Отчет Шт {date_str}.xls")
            main_blob = data_store.get_last_routes("main") or {}
            inc_blob = data_store.get_last_routes("increase") or {}
            excel_generator.generate_pcs_compare_report(
                report_path,
                main_blob.get("filteredRoutes") or main_blob.get("routes") or [],
                inc_blob.get("filteredRoutes") or inc_blob.get("routes") or [],
                data_store.get_ref("products") or [],
            )
            QMessageBox.information(self, "Готово", f"Файл создан:\n{save_path}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при создании файла:\n{e}")

    def _on_generate_all(self):
        if not self._dept_groups:
            QMessageBox.warning(self, "Нет данных",
                                "Нет отделов с маршрутами для генерации.")
            return

        save_dir = self.app_state.get("saveDir") or data_store.get_desktop_path()
        chosen_dir = QFileDialog.getExistingDirectory(
            self, "Выберите папку для сохранения файлов", save_dir
        )
        if not chosen_dir:
            return
        self.app_state["saveDir"] = chosen_dir

        prod_map  = data_store.get_products_map()
        templates = data_store.get_ref("templates") or []
        file_type = self.app_state.get("fileType", "main")
        data_store.save_last_routes(
            file_type,
            self.app_state.get("routes", []),
            self.app_state.get("uniqueProducts", []),
            self.app_state.get("filteredRoutes", []),
            route_category=self.app_state.get("routeCategory"),
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

        self._gen_thread = QThread(self)
        self._gen_worker = DeptGenerateWorker(
            self._dept_groups, file_type, chosen_dir, prod_map, templates, sort_asc,
            all_routes=all_routes, general_path=general_path, date_str=date_str
        )
        self._gen_worker.moveToThread(self._gen_thread)
        self._gen_thread.started.connect(self._gen_worker.run)
        self._gen_worker.finished.connect(self._on_gen_done)
        self._gen_worker.error.connect(self._on_gen_error)
        self._gen_worker.finished.connect(self._gen_thread.quit)
        self._gen_worker.error.connect(self._gen_thread.quit)
        self._gen_thread.start()

    def _on_gen_done(self, created: list):
        self.progress.setVisible(False)
        self.btn_generate_all.setEnabled(True)
        # Отчёт по Шт в корне папки дня.
        try:
            if created:
                any_path = created[0]
                day_dir = os.path.dirname(os.path.dirname(any_path))
                date_str = self.app_state.get("routesDate") or excel_generator.get_routes_date_str()
                report_path = os.path.join(day_dir, f"Отчет Шт {date_str}.xls")
                main_blob = data_store.get_last_routes("main") or {}
                inc_blob = data_store.get_last_routes("increase") or {}
                excel_generator.generate_pcs_compare_report(
                    report_path,
                    main_blob.get("filteredRoutes") or main_blob.get("routes") or [],
                    inc_blob.get("filteredRoutes") or inc_blob.get("routes") or [],
                    data_store.get_ref("products") or [],
                )
        except Exception:
            pass
        QMessageBox.information(
            self, "Готово",
            f"Создано файлов: {len(created)}\n\n" +
            "\n".join(os.path.basename(p) for p in created[:10]) +
            ("\n..." if len(created) > 10 else "")
        )

    def _on_gen_error(self, msg: str):
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
        if not any(p.get("labelTemplatePath") for p in products_ref):
            QMessageBox.information(
                self, "Нет шаблонов",
                "Откройте «Настройки этикеток» (страница Этикетки) и выберите шаблон XLS для продуктов."
            )
            return
        base_dir = self.app_state.get("saveDir") or data_store.get_desktop_path()
        tomorrow = date.today() + timedelta(days=1)
        folder_name = f"Этикетки на {tomorrow:%d.%m.%Y}"
        out_dir = os.path.join(base_dir, folder_name)
        os.makedirs(out_dir, exist_ok=True)
        file_type = self.app_state.get("fileType", "main")
        departments_ref = data_store.get_ref("departments") or []
        self._labels_out_dir = out_dir
        self.btn_labels.setEnabled(False)
        self.progress.setVisible(True)
        self._labels_thread = QThread(self)
        self._labels_worker = DeptTemplateLabelsWorker(
            routes, out_dir, file_type, products_ref, departments_ref
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
        self.progress.setVisible(False)
        self.btn_labels.setEnabled(True)
        QMessageBox.critical(self, "Ошибка", msg)

