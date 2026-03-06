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
from ui.widgets import make_combo_searchable


# ─────────────────────────── Worker ───────────────────────────────────────

class DeptGenerateWorker(QObject):
    finished = pyqtSignal(list)
    error    = pyqtSignal(str)

    def __init__(self, dept_groups: list, file_type: str, save_dir: str,
                 prod_map: dict, templates: list, sort_asc: bool = False):
        super().__init__()
        self.dept_groups = dept_groups
        self.file_type   = file_type
        self.save_dir    = save_dir
        self.prod_map    = prod_map
        self.templates   = templates
        self.sort_asc    = sort_asc

    def run(self):
        try:
            created = excel_generator.generate_dept_files(
                self.dept_groups, self.file_type,
                self.save_dir, self.prod_map, self.templates,
                self.sort_asc
            )
            self.finished.emit(created)
        except Exception as e:
            self.error.emit(str(e))


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
        self._table_font_size = 13
        self._dept_tables: list[QTableWidget] = []
        self._display_mode = _DISPLAY_ADDR_ONLY  # по умолчанию только № маршрута и адрес
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
        inner = QVBoxLayout(content)
        inner.setContentsMargins(28, 20, 28, 20)
        inner.setSpacing(16)

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
        self.lbl_font_size = QLabel("Размер текста: 13")
        self.lbl_font_size.setObjectName("badge")
        h_row.addWidget(self.lbl_font_size)
        inner.addLayout(h_row)

        # Фильтр по отделу и режим отображения
        filter_row = QHBoxLayout()
        lbl_filter = QLabel("Показать отдел/подотдел:")
        filter_row.addWidget(lbl_filter)
        self.combo_dept_filter = QComboBox()
        self.combo_dept_filter.setMinimumWidth(220)
        self.combo_dept_filter.currentIndexChanged.connect(self._on_filter_changed)
        filter_row.addWidget(self.combo_dept_filter)
        filter_row.addSpacing(16)
        filter_row.addWidget(QLabel("Показать:"))
        self.combo_display = QComboBox()
        self.combo_display.setMinimumWidth(140)
        self.combo_display.addItem("Только № и адрес", _DISPLAY_ADDR_ONLY)
        self.combo_display.addItem("Полностью", _DISPLAY_FULL)
        self.combo_display.setCurrentIndex(0)
        make_combo_searchable(self.combo_dept_filter)
        make_combo_searchable(self.combo_display)
        self.combo_display.currentIndexChanged.connect(self._on_display_changed)
        filter_row.addWidget(self.combo_display)
        filter_row.addStretch()
        inner.addLayout(filter_row)

        self.tabs = QTabWidget()
        inner.addWidget(self.tabs)

        bottom_row = QHBoxLayout()
        bottom_row.addStretch()

        self.btn_generate_all = QPushButton("Создать файлы для всех отделов")
        self.btn_generate_all.setObjectName("btnPrimary")
        self.btn_generate_all.clicked.connect(self._on_generate_all)
        bottom_row.addWidget(self.btn_generate_all)

        self.btn_labels = QPushButton("Этикетки из шаблонов (XLS)")
        self.btn_labels.setObjectName("btnSecondary")
        self.btn_labels.clicked.connect(self._on_labels_from_templates)
        bottom_row.addWidget(self.btn_labels)

        self.btn_clear = QPushButton("Очистить маршруты")
        self.btn_clear.setObjectName("btnDanger")
        self.btn_clear.clicked.connect(self._on_clear_routes)
        bottom_row.addWidget(self.btn_clear)

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
        assigned = {p["name"] for p in products if p.get("deptKey")}
        unassigned: set[str] = set()
        for r in routes:
            for p in r.get("products", []):
                if p["name"] not in assigned:
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
                        "routeNum": r["routeNum"],
                        "address":  r["address"],
                        "products": dept_prods,
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

    # ─────────────────────────── Рендер ───────────────────────────────────

    def _on_filter_changed(self, _index: int):
        """Переключает вкладку при изменении фильтра."""
        key = self.combo_dept_filter.currentData()
        if key is None:
            return  # «Все отделы» — ничего не делаем, все вкладки видны
        # Найти вкладку с нужным ключом
        for i, g in enumerate(self._dept_groups):
            if g["key"] == key:
                self.tabs.setCurrentIndex(i)
                break

    def refresh(self):
        file_type = self.app_state.get("fileType", "main")
        self.lbl_title.setText(
            f"Маршруты по отделам — "
            f"{'Основной' if file_type == 'main' else 'Увеличение (Довоз)'}"
        )

        # ── Проверка: есть ли продукты без отдела ──────────────────────────
        unassigned = self._check_unassigned_products()
        if unassigned:
            names_str = "\n  • ".join(unassigned[:15])
            if len(unassigned) > 15:
                names_str += f"\n  ... и ещё {len(unassigned)-15}"
            reply = QMessageBox.warning(
                self,
                "Продукты без отдела",
                f"Следующие продукты не привязаны к отделу:\n\n  • {names_str}\n\n"
                f"Пока все продукты не будут привязаны, страница маршрутов по отделам "
                f"будет работать некорректно.\n\n"
                f"Открыть окно «Отделы и продукты» для привязки?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                dept_mod.open_modal(self, self.app_state)
                # После закрытия окна — перестроить группы
                unassigned2 = self._check_unassigned_products()
                if unassigned2:
                    # Ещё остались непривязанные — заблокировать страницу
                    self._show_blocked(unassigned2)
                    return
            else:
                self._show_blocked(unassigned)
                return

        self._dept_groups = self._build_dept_groups()

        # Обновляем фильтр
        self.combo_dept_filter.blockSignals(True)
        self.combo_dept_filter.clear()
        self.combo_dept_filter.addItem("Все отделы", None)
        for g in self._dept_groups:
            prefix = "  └ " if g.get("is_subdept") else "• "
            self.combo_dept_filter.addItem(f"{prefix}{g['name']}", g["key"])
        self.combo_dept_filter.blockSignals(False)

        self.tabs.clear()
        self._dept_tables.clear()

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

    def _show_blocked(self, unassigned: list):
        """Показывает заглушку, пока есть непривязанные продукты."""
        self._dept_groups = []
        self.combo_dept_filter.clear()
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
        btn_open.setObjectName("btnPrimary")
        btn_open.clicked.connect(lambda: self._open_depts_and_retry())
        btn_retry = QPushButton("Проверить снова")
        btn_retry.setObjectName("btnSecondary")
        btn_retry.clicked.connect(self.refresh)
        btn_row.addWidget(btn_open, alignment=Qt.AlignmentFlag.AlignCenter)
        btn_row.addWidget(btn_retry, alignment=Qt.AlignmentFlag.AlignCenter)
        blay.addLayout(btn_row)
        self.tabs.addTab(blocked, "Заблокировано")

    def _open_depts_and_retry(self):
        dept_mod.open_modal(self, self.app_state)
        self.refresh()

    def _make_dept_tab(self, group: dict, prod_map: dict) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(8, 8, 8, 8)
        lay.setSpacing(8)

        lbl = QLabel(f"Маршрутов: {len(group['routes'])}")
        lbl.setObjectName("hintLabel")
        lay.addWidget(lbl)

        table = QTableWidget()
        table.setColumnCount(5)
        table.setHorizontalHeaderLabels(["№ маршрута", "Адрес / Продукт", "Ед. изм.", "Кол-во", "Шт"])
        hdr = table.horizontalHeader()
        hdr.setMinimumSectionSize(90)
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.verticalHeader().setVisible(False)
        table.setAlternatingRowColors(True)
        table.setFont(QFont("", self._table_font_size))
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
                self.lbl_font_size.setText(f"Размер текста: {self._table_font_size}")
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
        sort_asc = self.app_state.get("sortAsc", False)
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

            # Строка маршрута: номер | адрес | пусто | пусто | пусто
            table.setItem(current_row, 0, _make_item(route_num, bold=True))
            table.setItem(current_row, 1, _make_item(address,   bold=True, bg=gray_bg))
            table.setItem(current_row, 2, _make_item(""))
            table.setItem(current_row, 3, _make_item(""))
            table.setItem(current_row, 4, _make_item(""))
            current_row += 1

            if not only_addr:
                # Строки продуктов: пусто | название | ед.изм. | кол-во | шт
                for prod in products:
                    table.setItem(current_row, 0, _make_item(""))
                    table.setItem(current_row, 1, _make_item(f"  {prod.get('name', '')}"))
                    table.setItem(current_row, 2, _make_item(prod.get("unit", "")))
                    qty_str = self._fmt_qty(prod, prod_map)
                    pcs_str = self._fmt_pcs(prod, prod_map, route)
                    table.setItem(current_row, 3, _make_item(qty_str, align_right=True))
                    table.setItem(current_row, 4, _make_item(pcs_str, align_right=True))
                    current_row += 1

        table.setUpdatesEnabled(True)
        table.resizeColumnsToContents()

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

        route_cat = route.get("routeCategory") or "ШК"
        round_up = (
            ps.get("roundUpСД") if "roundUpСД" in ps else ps.get("roundUp", True)
            if route_cat == "СД"
            else ps.get("roundUpШК") if "roundUpШК" in ps else ps.get("roundUp", True)
        )
        pcs = excel_generator.calc_pcs(
            display_qty, float(ps.get("pcsPerUnit", 1)), bool(round_up)
        )
        return str(pcs)

    # ─────────────────────────── Генерация ────────────────────────────────

    def _on_save_single(self, group: dict):
        tomorrow = datetime.now() + timedelta(days=1)
        date_str  = tomorrow.strftime("%d.%m.%Y")
        file_type = self.app_state.get("fileType", "main")
        suffix    = "УВ" if file_type == "increase" else "ОСН"

        import re
        safe_name = re.sub(r'[\\/:*?"<>|]', "_", group["name"])
        default_name = f"Маршруты {safe_name} {date_str} {suffix}.xls"

        save_dir = self.app_state.get("saveDir") or data_store.get_desktop_path()
        save_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить файл",
            os.path.join(save_dir, default_name),
            "Excel 97-2003 (*.xls)"
        )
        if not save_path:
            return
        if not save_path.lower().endswith(".xls"):
            save_path += ".xls"

        try:
            prod_map  = data_store.get_products_map()
            templates = data_store.get_ref("templates") or []
            sort_asc  = self.app_state.get("sortAsc", False)
            excel_generator.generate_single_dept_file(
                group, file_type, save_path, prod_map, templates, sort_asc
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

        prod_map  = data_store.get_products_map()
        templates = data_store.get_ref("templates") or []
        file_type = self.app_state.get("fileType", "main")
        sort_asc  = self.app_state.get("sortAsc", False)

        self.btn_generate_all.setEnabled(False)
        self.progress.setVisible(True)

        self._gen_thread = QThread(self)
        self._gen_worker = DeptGenerateWorker(
            self._dept_groups, file_type, chosen_dir, prod_map, templates, sort_asc
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
        try:
            created = excel_generator.generate_labels_from_templates(
                routes, out_dir, file_type, products_ref, departments_ref
            )
            if created:
                set_status = self.app_state.get("set_status")
                if callable(set_status):
                    set_status(f"Создано этикеток: {len(created)}")
                QMessageBox.information(self, "Готово", f"Создано файлов: {len(created)}\n\n{out_dir}")
            else:
                QMessageBox.information(self, "Нет файлов", "Нет этикеток для создания.")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))

    def _on_clear_routes(self):
        reply = QMessageBox.question(
            self, "Очистить маршруты",
            "Удалить все загруженные маршруты и последние сохранённые данные?\n"
            "После этого можно загрузить новые файлы.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            data_store.clear_last_routes()
            self.app_state.update({
                "filePaths": [], "routes": [], "uniqueProducts": [],
                "filteredRoutes": [], "routeCategory": "ШК",
            })
            self.go_clear_routes.emit()
