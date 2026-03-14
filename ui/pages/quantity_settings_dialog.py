"""
quantity_settings_dialog.py — Настройки Количества.

Окно: список продуктов (привязанных к отделам), без канонического названия и без удаления.
По двойному клику по продукту открывается настройка в Шт (Показывать Шт, Кол-во в 1 шт,
Хвостик ШК, Хвостик СД) справа; повторный двойной клик по тому же продукту скрывает панель.
Ниже — блок «Округление по учреждениям» (выбор учреждений с округлением Шт вверх).
"""
from __future__ import annotations

import copy
from typing import Iterable

from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QFont, QColor, QBrush
from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QListWidget,
    QListWidgetItem,
    QDoubleSpinBox,
    QLineEdit,
    QWidget,
    QScrollArea,
    QFrame,
    QGroupBox,
    QTabWidget,
)

from core import data_store, excel_generator
from ui.widgets import ToggleSwitch


def _extract_institution_list(routes: Iterable[dict]) -> list[str]:
    """Список уникальных учреждений по маршрутам (3–4 цифры или полный адрес)."""
    return data_store.get_institution_list_from_routes(routes)


def _tail_value_to_percent(tail_value: float | None, pcs_per_unit: float) -> float:
    """Переводит абсолютный порог остатка в процент от 1 шт."""
    if pcs_per_unit <= 0:
        return 0.0
    if tail_value is None:
        return 0.0
    return max(0.0, min(100.0, (float(tail_value) / float(pcs_per_unit)) * 100.0))


def _percent_to_tail_value(percent_value: float, pcs_per_unit: float) -> float:
    """Переводит процент от 1 шт в абсолютный порог остатка."""
    if pcs_per_unit <= 0:
        return 0.0
    pct = max(0.0, min(100.0, float(percent_value)))
    return float(pcs_per_unit) * (pct / 100.0)


def _format_pcs_total(total: float | int | None) -> str:
    if total is None:
        return ""
    if abs(float(total) - round(float(total))) < 1e-9:
        return f"{int(round(float(total)))} шт"
    return f"{float(total):.1f} шт"


def _calc_product_pcs_totals(app_state: dict) -> dict[str, float]:
    """Считает общую сумму шт по каждому продукту в текущих маршрутах."""
    routes = app_state.get("filteredRoutes") or app_state.get("routes") or []
    if not routes:
        return {}
    routes_copy = copy.deepcopy(routes)
    products_ref = data_store.get_ref("products") or []
    prod_map = {p.get("name"): dict(p) for p in products_ref if p.get("name")}
    excel_generator._apply_pcs(routes_copy, prod_map)
    totals: dict[str, float] = {}
    for route in routes_copy:
        for prod in route.get("products", []):
            pcs = prod.get("pcs")
            if pcs is None:
                continue
            try:
                totals[prod.get("name", "")] = totals.get(prod.get("name", ""), 0.0) + float(pcs)
            except (TypeError, ValueError):
                continue
    return totals


# ─────────────────────────── Панель настроек одного продукта ─────────────────

class ProductPcsPanel(QFrame):
    """Панель настроек 1 шт для одного продукта."""

    def __init__(self, product_name: str, on_changed=None, parent=None):
        super().__init__(parent)
        self._product_name = product_name
        self._on_changed_cb = on_changed
        self._loading = False
        self.setObjectName("card")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 12, 16, 12)
        lay.setSpacing(10)

        self.lbl_total_pcs = QLabel("")
        self.lbl_total_pcs.setObjectName("hintLabel")
        self.lbl_total_pcs.setWordWrap(True)
        self.lbl_total_pcs.setVisible(False)
        lay.addWidget(self.lbl_total_pcs)

        # Показывать Шт
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("Показывать Шт:"))
        self.chk_show = ToggleSwitch()
        self.chk_show.stateChanged.connect(self._on_show_changed)
        row1.addWidget(self.chk_show)
        row1.addStretch()
        lay.addLayout(row1)

        # Кол-во в 1 шт
        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Кол-во в 1 шт:"))
        self.spin_pcs = QDoubleSpinBox()
        self.spin_pcs.setRange(0.001, 99999.0)
        self.spin_pcs.setDecimals(3)
        self.spin_pcs.setSingleStep(0.1)
        self.spin_pcs.valueChanged.connect(self._on_pcs_changed)
        row2.addWidget(self.spin_pcs)
        row2.addStretch()
        lay.addLayout(row2)

        self.lbl_hint = QLabel(
            "Пороги округления Шт для школ и садов задаются выше на уровне отдела или подотдела."
        )
        self.lbl_hint.setObjectName("hintLabel")
        self.lbl_hint.setWordWrap(True)
        lay.addWidget(self.lbl_hint)

        self._load()

    def _load(self) -> None:
        self._loading = True
        products = data_store.get_ref("products") or []
        prod = next((p for p in products if p.get("name") == self._product_name), None)
        if not prod:
            self.chk_show.setEnabled(False)
            self.spin_pcs.setEnabled(False)
            self._loading = False
            return
        unit = (prod.get("unit") or "").strip().lower()
        if unit == "шт":
            self.chk_show.setEnabled(False)
            self.spin_pcs.setEnabled(False)
            self._loading = False
            return
        show_pcs = prod.get("showPcs", False)
        self.chk_show.setChecked(show_pcs)
        pcs_per_unit = float(prod.get("pcsPerUnit", 1.0) or 1.0)
        self.spin_pcs.setValue(pcs_per_unit)
        self.spin_pcs.setEnabled(show_pcs)
        self._loading = False

    def set_total_pcs_text(self, text: str) -> None:
        self.lbl_total_pcs.setText(text)
        self.lbl_total_pcs.setVisible(bool(text))

    def _notify_changed(self) -> None:
        if callable(self._on_changed_cb):
            self._on_changed_cb(self._product_name)

    def _on_show_changed(self, state: int) -> None:
        if self._loading:
            return
        show = state == Qt.CheckState.Checked.value
        self.spin_pcs.setEnabled(show)
        data_store.update_product(self._product_name, showPcs=show)
        self._notify_changed()

    def _on_pcs_changed(self, val: float) -> None:
        if self._loading:
            return
        products = data_store.get_ref("products") or []
        prod = next((p for p in products if p.get("name") == self._product_name), {}) or {}
        old_pcs = float(prod.get("pcsPerUnit", 1.0) or 1.0)
        shk_percent = _tail_value_to_percent(
            prod.get("roundTailFromШК")
            if prod.get("roundTailFromШК") is not None
            else (0 if prod.get("roundUpШК", True) else old_pcs),
            old_pcs,
        )
        sd_percent = _tail_value_to_percent(
            prod.get("roundTailFromСД")
            if prod.get("roundTailFromСД") is not None
            else (0 if prod.get("roundUpСД", True) else old_pcs),
            old_pcs,
        )
        data_store.update_product(
            self._product_name,
            pcsPerUnit=val,
            roundTailFromШК=_percent_to_tail_value(shk_percent, val),
            roundTailFromСД=_percent_to_tail_value(sd_percent, val),
        )
        self._notify_changed()


# ─────────────────────────── Строка списка: название продукта ────────────────

class ProductRowWidget(QWidget):
    """Одна строка: название продукта (настройки открываются по двойному клику)."""

    ROW_HEIGHT = 70

    def __init__(self, product_name: str, parent=None):
        super().__init__(parent)
        self.product_name = product_name
        self.setMinimumHeight(self.ROW_HEIGHT)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(12, 12, 12, 12)
        text_col = QVBoxLayout()
        text_col.setSpacing(2)

        self.lbl_name = QLabel(product_name)
        self.lbl_name.setObjectName("cardTitle")
        self.lbl_name.setAlignment(Qt.AlignmentFlag.AlignVCenter)
        self.lbl_name.setWordWrap(True)
        self.lbl_name.setMinimumHeight(28)
        font = QFont()
        font.setPointSize(14)
        self.lbl_name.setFont(font)
        text_col.addWidget(self.lbl_name)

        self.lbl_total = QLabel("")
        self.lbl_total.setObjectName("hintLabel")
        self.lbl_total.setWordWrap(True)
        self.lbl_total.setVisible(False)
        text_col.addWidget(self.lbl_total)

        self.lbl_mode = QLabel("")
        self.lbl_mode.setObjectName("stepLabel")
        self.lbl_mode.setWordWrap(True)
        self.lbl_mode.setVisible(False)
        text_col.addWidget(self.lbl_mode)

        lay.addLayout(text_col, 1)

    def sizeHint(self) -> QSize:
        return QSize(260, self.ROW_HEIGHT)

    def set_total_pcs_text(self, text: str) -> None:
        self.lbl_total.setText(text)
        self.lbl_total.setVisible(bool(text))

    def set_mode_text(self, text: str) -> None:
        self.lbl_mode.setText(text)
        self.lbl_mode.setVisible(bool(text))


class DeptRoundPanel(QFrame):
    """Панель настройки округления Шт для выбранного контекста отдела."""

    def __init__(
        self,
        scope_id: str,
        scope_label: str,
        dept_keys: list[str],
        on_changed=None,
        parent=None,
    ):
        super().__init__(parent)
        self._scope_id = scope_id
        self._scope_label = scope_label
        self._dept_keys = [k for k in dept_keys if k]
        self._on_changed_cb = on_changed
        self._loading = False
        self.setObjectName("card")

        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 12, 16, 12)
        lay.setSpacing(10)

        self.lbl_title = QLabel(scope_label)
        self.lbl_title.setObjectName("sectionTitle")
        self.lbl_title.setWordWrap(True)
        lay.addWidget(self.lbl_title)

        self.lbl_meta = QLabel("")
        self.lbl_meta.setObjectName("hintLabel")
        self.lbl_meta.setWordWrap(True)
        lay.addWidget(self.lbl_meta)

        self.lbl_mixed = QLabel("")
        self.lbl_mixed.setObjectName("stepLabel")
        self.lbl_mixed.setWordWrap(True)
        self.lbl_mixed.setVisible(False)
        lay.addWidget(self.lbl_mixed)

        row_shk = QHBoxLayout()
        row_shk.addWidget(QLabel("Школы (% остатка для округления):"))
        self.spin_shk = QDoubleSpinBox()
        self.spin_shk.setRange(0, 100.0)
        self.spin_shk.setDecimals(1)
        self.spin_shk.setSuffix(" %")
        self.spin_shk.setSingleStep(5.0)
        self.spin_shk.valueChanged.connect(self._on_shk_changed)
        row_shk.addWidget(self.spin_shk)
        row_shk.addStretch()
        lay.addLayout(row_shk)

        row_sd = QHBoxLayout()
        row_sd.addWidget(QLabel("Сады (% остатка для округления):"))
        self.spin_sd = QDoubleSpinBox()
        self.spin_sd.setRange(0, 100.0)
        self.spin_sd.setDecimals(1)
        self.spin_sd.setSuffix(" %")
        self.spin_sd.setSingleStep(5.0)
        self.spin_sd.valueChanged.connect(self._on_sd_changed)
        row_sd.addWidget(self.spin_sd)
        row_sd.addStretch()
        lay.addLayout(row_sd)

        lay.addStretch()
        self._load()

    def _dept_products(self) -> list[dict]:
        products = data_store.get_ref("products") or []
        return [
            p for p in products
            if p.get("deptKey") in self._dept_keys
            and p.get("name")
            and (p.get("unit") or "").strip().lower() != "шт"
        ]

    def _product_percent(self, prod: dict, route_cat: str) -> float:
        pcs_per_unit = float(prod.get("pcsPerUnit", 1.0) or 1.0)
        if route_cat == "СД":
            tail_value = prod.get("roundTailFromСД")
            if tail_value is None:
                tail_value = 0 if prod.get("roundUpСД", True) else pcs_per_unit
        else:
            tail_value = prod.get("roundTailFromШК")
            if tail_value is None:
                tail_value = 0 if prod.get("roundUpШК", True) else pcs_per_unit
        return round(_tail_value_to_percent(tail_value, pcs_per_unit), 1)

    def _load(self) -> None:
        self._loading = True
        products = self._dept_products()
        total = len(products)
        enabled = sum(1 for p in products if p.get("showPcs"))
        self.lbl_title.setText(self._scope_label)
        self.lbl_meta.setText(
            f"Товаров в текущем блоке: {total}. С включённым отображением Шт: {enabled}."
        )
        shk_values = sorted({self._product_percent(prod, "ШК") for prod in products}) or [0.0]
        sd_values = sorted({self._product_percent(prod, "СД") for prod in products}) or [0.0]
        self.spin_shk.setValue(shk_values[0])
        self.spin_sd.setValue(sd_values[0])
        mixed_parts = []
        if len(shk_values) > 1:
            mixed_parts.append("Школы")
        if len(sd_values) > 1:
            mixed_parts.append("Сады")
        if mixed_parts:
            self.lbl_mixed.setText(
                "Сейчас у товаров отдела разные значения: "
                + ", ".join(mixed_parts)
                + ". При изменении здесь они будут выровнены."
            )
            self.lbl_mixed.setVisible(True)
        else:
            self.lbl_mixed.setVisible(False)
            self.lbl_mixed.setText("")
        self._loading = False

    def _apply_percent(self, route_cat: str, percent_value: float) -> None:
        products = data_store.get("products") or []
        field = "roundTailFromСД" if route_cat == "СД" else "roundTailFromШК"
        changed = False
        for prod in products:
            if prod.get("deptKey") not in self._dept_keys:
                continue
            if not prod.get("name"):
                continue
            if (prod.get("unit") or "").strip().lower() == "шт":
                continue
            pcs_per_unit = float(prod.get("pcsPerUnit", 1.0) or 1.0)
            prod[field] = _percent_to_tail_value(percent_value, pcs_per_unit)
            changed = True
        if not changed:
            return
        data_store.set_key("products", products)
        self._load()
        if callable(self._on_changed_cb):
            self._on_changed_cb(self._scope_id)

    def _on_shk_changed(self, val: float) -> None:
        if self._loading:
            return
        self._apply_percent("ШК", val)

    def _on_sd_changed(self, val: float) -> None:
        if self._loading:
            return
        self._apply_percent("СД", val)


class InstitutionStatusRowWidget(QFrame):
    """Строка учреждения с заметным бейджем статуса."""

    ROW_HEIGHT = 72

    def __init__(self, code: str, parent=None):
        super().__init__(parent)
        self.code = code
        self._selected = False
        self.setObjectName("card")
        self.setMinimumHeight(self.ROW_HEIGHT)

        lay = QHBoxLayout(self)
        lay.setContentsMargins(12, 10, 12, 10)
        lay.setSpacing(10)

        text_col = QVBoxLayout()
        text_col.setSpacing(4)

        self.lbl_code = QLabel(code)
        self.lbl_code.setObjectName("cardTitle")
        self.lbl_code.setWordWrap(True)
        text_col.addWidget(self.lbl_code)

        self.lbl_meta = QLabel()
        self.lbl_meta.setObjectName("hintLabel")
        self.lbl_meta.setWordWrap(True)
        text_col.addWidget(self.lbl_meta)

        lay.addLayout(text_col, 1)

        self.lbl_badge = QLabel()
        self.lbl_badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_badge.setMinimumWidth(108)
        lay.addWidget(self.lbl_badge)

    def sizeHint(self) -> QSize:
        return QSize(320, self.ROW_HEIGHT)

    def update_content(self, status: str, color: QColor, total: int, active: int) -> None:
        self.lbl_meta.setText(f"Адресов: {total}. Активных: {active}.")
        self.lbl_badge.setText(status.upper())
        self.lbl_badge.setStyleSheet(
            f"color: {color.name()};"
            f"background: rgba({color.red()}, {color.green()}, {color.blue()}, 0.14);"
            f"border: 1px solid {color.name()};"
            "border-radius: 10px;"
            "padding: 6px 10px;"
            "font-weight: 700;"
        )
        self._apply_selected_style()

    def set_selected(self, selected: bool) -> None:
        self._selected = selected
        self._apply_selected_style()

    def _apply_selected_style(self) -> None:
        if self._selected:
            self.setStyleSheet("QFrame#card { border: 2px solid #2563EB; background-color: #F5F9FF; }")
        else:
            self.setStyleSheet("")


class InstitutionsSettingsWidget(QFrame):
    """Встроенный блок настройки округления по учреждениям."""

    def __init__(self, app_state: dict, parent: QWidget | None = None):
        super().__init__(parent)
        self._app_state = app_state
        routes = app_state.get("routes") or app_state.get("filteredRoutes") or []
        self._inst_addresses = data_store.get_institution_addresses_map(routes)
        self._selected_codes: set[str] = set(
            data_store.get_setting("alwaysRoundUpInstitutions") or []
        )
        self._excluded_addresses: set[str] = set(
            data_store.get_setting("excludeRoundUpAddresses") or []
        )
        self._current_code: str | None = None

        self._build_ui()
        self._populate_institutions()
        self._populate_dept_percents()

    def _build_ui(self) -> None:
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(12)

        hint = QLabel(
            "Настройка работает прямо в этом окне: выберите учреждение слева и при необходимости "
            "исключите отдельные адреса справа."
        )
        hint.setObjectName("stepLabel")
        hint.setWordWrap(True)
        lay.addWidget(hint)

        self.lbl_no_routes = QLabel("Нет учреждений в текущих маршрутах.")
        self.lbl_no_routes.setObjectName("hintLabel")
        self.lbl_no_routes.setVisible(False)
        lay.addWidget(self.lbl_no_routes)

        body = QHBoxLayout()
        body.setSpacing(16)

        left_card = QFrame()
        left_card.setObjectName("card")
        left_lay = QVBoxLayout(left_card)
        left_lay.setContentsMargins(16, 16, 16, 16)
        left_lay.setSpacing(10)

        left_lay.addWidget(QLabel("Учреждения"))

        self.search_inst = QLineEdit()
        self.search_inst.setPlaceholderText("Поиск учреждения или адреса")
        self.search_inst.textChanged.connect(self._apply_inst_filter)
        left_lay.addWidget(self.search_inst)

        bulk_row = QHBoxLayout()
        self.btn_enable_all_inst = QPushButton("Включить все")
        self.btn_enable_all_inst.setObjectName("btnSecondary")
        self.btn_enable_all_inst.clicked.connect(lambda: self._set_all_institutions(True))
        bulk_row.addWidget(self.btn_enable_all_inst)
        self.btn_disable_all_inst = QPushButton("Выключить все")
        self.btn_disable_all_inst.setObjectName("btnSecondary")
        self.btn_disable_all_inst.clicked.connect(lambda: self._set_all_institutions(False))
        bulk_row.addWidget(self.btn_disable_all_inst)
        left_lay.addLayout(bulk_row)

        self.lbl_inst_summary = QLabel()
        self.lbl_inst_summary.setObjectName("hintLabel")
        self.lbl_inst_summary.setWordWrap(True)
        left_lay.addWidget(self.lbl_inst_summary)

        self.inst_list = QListWidget()
        self.inst_list.setAlternatingRowColors(True)
        self.inst_list.currentItemChanged.connect(self._on_institution_selected)
        left_lay.addWidget(self.inst_list, 1)

        body.addWidget(left_card, 1)

        self.tabs = QTabWidget()

        tab_addresses = QWidget()
        right_lay = QVBoxLayout(tab_addresses)
        right_lay.setContentsMargins(16, 16, 16, 16)
        right_lay.setSpacing(12)

        self.lbl_current_inst = QLabel("Выберите учреждение слева")
        self.lbl_current_inst.setObjectName("sectionTitle")
        self.lbl_current_inst.setWordWrap(True)
        right_lay.addWidget(self.lbl_current_inst)

        self.lbl_current_status = QLabel("")
        self.lbl_current_status.setObjectName("hintLabel")
        self.lbl_current_status.setWordWrap(True)
        right_lay.addWidget(self.lbl_current_status)

        self.lbl_current_help = QLabel(
            "Для выбранного учреждения можно включить округление и настроить исключения по адресам."
        )
        self.lbl_current_help.setObjectName("stepLabel")
        self.lbl_current_help.setWordWrap(True)
        right_lay.addWidget(self.lbl_current_help)

        row_inst_toggle = QHBoxLayout()
        row_inst_toggle.addWidget(QLabel("Округлять для всего учреждения:"))
        self.tog_current_inst = ToggleSwitch()
        self.tog_current_inst.stateChanged.connect(self._on_current_inst_toggle)
        row_inst_toggle.addWidget(self.tog_current_inst)
        row_inst_toggle.addStretch()
        right_lay.addLayout(row_inst_toggle)

        actions_row = QHBoxLayout()
        self.btn_include_all_addr = QPushButton("Включить все адреса")
        self.btn_include_all_addr.setObjectName("btnSecondary")
        self.btn_include_all_addr.clicked.connect(lambda: self._set_all_addresses_for_current(True))
        actions_row.addWidget(self.btn_include_all_addr)
        self.btn_exclude_all_addr = QPushButton("Исключить все адреса")
        self.btn_exclude_all_addr.setObjectName("btnSecondary")
        self.btn_exclude_all_addr.clicked.connect(lambda: self._set_all_addresses_for_current(False))
        actions_row.addWidget(self.btn_exclude_all_addr)
        actions_row.addStretch()
        right_lay.addLayout(actions_row)

        addr_filter_row = QHBoxLayout()
        addr_filter_row.addWidget(QLabel("Поиск адреса:"))
        self.search_addr = QLineEdit()
        self.search_addr.setPlaceholderText("Введите часть адреса")
        self.search_addr.textChanged.connect(lambda _t: self._refresh_selected_panel())
        addr_filter_row.addWidget(self.search_addr, 1)
        addr_filter_row.addWidget(QLabel("Только исключённые:"))
        self.tog_only_excluded = ToggleSwitch()
        self.tog_only_excluded.stateChanged.connect(lambda _s: self._refresh_selected_panel())
        addr_filter_row.addWidget(self.tog_only_excluded)
        addr_filter_row.addStretch()
        right_lay.addLayout(addr_filter_row)

        self.lbl_addr_summary = QLabel()
        self.lbl_addr_summary.setObjectName("hintLabel")
        self.lbl_addr_summary.setWordWrap(True)
        right_lay.addWidget(self.lbl_addr_summary)

        self.addr_scroll = QScrollArea()
        self.addr_scroll.setWidgetResizable(True)
        self.addr_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.addr_container = QWidget()
        self.addr_lay = QVBoxLayout(self.addr_container)
        self.addr_lay.setContentsMargins(0, 0, 0, 0)
        self.addr_lay.setSpacing(8)
        self.addr_scroll.setWidget(self.addr_container)
        right_lay.addWidget(self.addr_scroll, 1)

        tab_percent = QWidget()
        percent_lay = QVBoxLayout(tab_percent)
        percent_lay.setContentsMargins(16, 16, 16, 16)
        percent_lay.setSpacing(12)

        grp_dept = QGroupBox("Округление при остатке ≥ % от 1 шт")
        grp_lay = QVBoxLayout(grp_dept)
        hint_dept = QLabel("Для каждого отдела — свой порог (если не задан, используется общий).")
        hint_dept.setObjectName("stepLabel")
        hint_dept.setWordWrap(True)
        grp_lay.addWidget(hint_dept)

        self.spin_default = QDoubleSpinBox()
        self.spin_default.setRange(0, 100)
        self.spin_default.setDecimals(1)
        self.spin_default.setSuffix(" %")
        self.spin_default.setSingleStep(5)
        self.spin_default.setValue(float(data_store.get_setting("roundUpInstitutionPercent") or 20))
        self.spin_default.valueChanged.connect(self._on_default_pct_changed)
        def_row = QHBoxLayout()
        def_row.addWidget(QLabel("Общий % (по умолчанию):"))
        def_row.addWidget(self.spin_default)
        def_row.addStretch()
        grp_lay.addLayout(def_row)

        self.dept_percent_widget = QWidget()
        self.dept_percent_lay = QVBoxLayout(self.dept_percent_widget)
        self.dept_percent_lay.setContentsMargins(0, 8, 0, 0)
        grp_lay.addWidget(self.dept_percent_widget)
        percent_lay.addWidget(grp_dept)
        percent_lay.addStretch()

        self.tabs.addTab(tab_addresses, "Адреса")
        self.tabs.addTab(tab_percent, "% по отделам")
        body.addWidget(self.tabs, 2)
        lay.addLayout(body, 1)

    def _status_info(self, code: str) -> tuple[str, QColor]:
        addresses = self._inst_addresses.get(code, [])
        total = len(addresses)
        active = self._effective_active_count(code)
        enabled = code in self._selected_codes
        if enabled and total > 0 and active == total:
            return "Включено", QColor("#22C55E")
        if enabled and active > 0:
            return "Частично", QColor("#F59E0B")
        return "Выключено", QColor("#9CA3AF")

    def _manual_excluded_count(self, code: str) -> int:
        addresses = self._inst_addresses.get(code, [])
        return sum(1 for addr in addresses if addr in self._excluded_addresses)

    def _effective_active_count(self, code: str) -> int:
        if code not in self._selected_codes:
            return 0
        addresses = self._inst_addresses.get(code, [])
        return sum(1 for addr in addresses if addr not in self._excluded_addresses)

    def _status_badge_text(self, status: str) -> str:
        if status == "Включено":
            return "Статус: ВКЛЮЧЕНО"
        if status == "Частично":
            return "Статус: ЧАСТИЧНО"
        return "Статус: ВЫКЛЮЧЕНО"

    def _refresh_inst_summary(self) -> None:
        all_codes = list(self._inst_addresses)
        enabled_count = sum(1 for code in all_codes if self._status_info(code)[0] == "Включено")
        partial_count = sum(1 for code in all_codes if self._status_info(code)[0] == "Частично")
        self.lbl_inst_summary.setText(
            f"Всего учреждений: {len(all_codes)}. Включено: {enabled_count}. Частично: {partial_count}."
        )

    def _populate_institutions(self) -> None:
        self.inst_list.clear()
        all_codes = sorted(self._inst_addresses)
        self.lbl_no_routes.setVisible(not all_codes)
        self.inst_list.setVisible(bool(all_codes))
        self.btn_enable_all_inst.setVisible(bool(all_codes))
        self.btn_disable_all_inst.setVisible(bool(all_codes))
        self.search_inst.setVisible(bool(all_codes))
        for code in all_codes:
            addresses = self._inst_addresses.get(code, [])
            active_addrs = sum(1 for addr in addresses if addr not in self._excluded_addresses)
            item = QListWidgetItem()
            item.setData(Qt.ItemDataRole.UserRole, code)
            item.setToolTip("\n".join(addresses[:10]))
            self.inst_list.addItem(item)
            row = InstitutionStatusRowWidget(code, self.inst_list)
            self.inst_list.setItemWidget(item, row)
            item.setSizeHint(row.sizeHint())
            self._update_inst_item(item)
        self._refresh_inst_summary()
        if all_codes:
            row_to_select = 0
            if self._current_code:
                for i in range(self.inst_list.count()):
                    if self.inst_list.item(i).data(Qt.ItemDataRole.UserRole) == self._current_code:
                        row_to_select = i
                        break
            self.inst_list.setCurrentRow(row_to_select)
        else:
            self._current_code = None
            self._refresh_selected_panel()

    def _update_inst_item(self, item: QListWidgetItem) -> None:
        code = item.data(Qt.ItemDataRole.UserRole)
        addresses = self._inst_addresses.get(code, [])
        active_addrs = self._effective_active_count(code)
        status, color = self._status_info(code)
        item.setForeground(QBrush(color))
        row = self.inst_list.itemWidget(item)
        if isinstance(row, InstitutionStatusRowWidget):
            row.update_content(status, color, len(addresses), active_addrs)

    def _clear_layout_widgets(self, layout: QVBoxLayout) -> None:
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            child_layout = item.layout()
            if widget is not None:
                widget.deleteLater()
            elif child_layout is not None:
                while child_layout.count():
                    child = child_layout.takeAt(0)
                    if child.widget():
                        child.widget().deleteLater()

    def _refresh_selected_panel(self) -> None:
        self._clear_layout_widgets(self.addr_lay)
        code = self._current_code
        has_code = bool(code)
        self.tog_current_inst.setEnabled(has_code)
        self.btn_include_all_addr.setEnabled(has_code)
        self.btn_exclude_all_addr.setEnabled(has_code)
        self.search_addr.setEnabled(has_code)
        self.tog_only_excluded.setEnabled(has_code)
        if not code:
            self.lbl_current_inst.setText("Выберите учреждение слева")
            self.lbl_current_inst.setStyleSheet("")
            self.lbl_current_status.setText("")
            self.lbl_current_help.setText(
                "Для выбранного учреждения можно включить округление и настроить исключения по адресам."
            )
            self.lbl_addr_summary.setText("")
            return

        addresses = self._inst_addresses.get(code, [])
        enabled = code in self._selected_codes
        active_addrs = self._effective_active_count(code)
        excluded_count = self._manual_excluded_count(code)
        addr_pattern = (self.search_addr.text() or "").strip().lower()
        only_excluded = self.tog_only_excluded.isChecked()

        self.lbl_current_inst.setText(code)
        status, color = self._status_info(code)
        self.lbl_current_inst.setStyleSheet(f"color: {color.name()};")
        self.lbl_current_status.setText(self._status_badge_text(status))
        self.lbl_current_status.setStyleSheet(
            f"color: {color.name()}; font-weight: 700; background: rgba({color.red()}, {color.green()}, {color.blue()}, 0.12); padding: 6px 10px; border: 1px solid {color.name()}; border-radius: 8px;"
        )
        if enabled:
            self.lbl_current_help.setText(
                "Если учреждение включено, округление будет применяться ко всем адресам ниже, "
                "кроме явно исключённых."
            )
        else:
            self.lbl_current_help.setText(
                "Учреждение выключено: округление сейчас не применяется ни к одному адресу ниже."
            )
        self.tog_current_inst.blockSignals(True)
        self.tog_current_inst.setChecked(enabled)
        self.tog_current_inst.blockSignals(False)
        self.lbl_addr_summary.setText(
            f"Адресов: {len(addresses)}. Активно: {active_addrs}. Исключено вручную: {excluded_count}."
        )

        if not addresses:
            empty = QLabel("В текущих маршрутах нет адресов для этого учреждения.")
            empty.setObjectName("hintLabel")
            empty.setWordWrap(True)
            self.addr_lay.addWidget(empty)
            self.addr_lay.addStretch()
            return

        shown_count = 0
        for addr in addresses:
            is_excluded = addr in self._excluded_addresses
            if only_excluded and not is_excluded:
                continue
            if addr_pattern and addr_pattern not in addr.lower():
                continue
            row = QFrame()
            row.setObjectName("card")
            row_lay = QHBoxLayout(row)
            row_lay.setContentsMargins(12, 10, 12, 10)
            row_lay.setSpacing(12)

            lbl = QLabel(addr)
            lbl.setWordWrap(True)
            lbl.setObjectName("stepLabel")
            row_lay.addWidget(lbl, 1)

            tog = ToggleSwitch()
            tog.setChecked(enabled and not is_excluded)
            tog.setEnabled(enabled)
            tog.stateChanged.connect(lambda state, a=addr: self._on_addr_toggle(a, state))
            row_lay.addWidget(tog)
            self.addr_lay.addWidget(row)
            shown_count += 1
        if shown_count == 0:
            empty = QLabel("Нет адресов по текущему фильтру.")
            empty.setObjectName("hintLabel")
            empty.setWordWrap(True)
            self.addr_lay.addWidget(empty)
        self.addr_lay.addStretch()

    def _apply_inst_filter(self, text: str) -> None:
        pattern = (text or "").strip().lower()
        first_visible = None
        for i in range(self.inst_list.count()):
            item = self.inst_list.item(i)
            code = str(item.data(Qt.ItemDataRole.UserRole) or "")
            addresses = self._inst_addresses.get(code, [])
            haystack = " ".join([code, *addresses]).lower()
            visible = not pattern or pattern in haystack
            item.setHidden(not visible)
            if visible and first_visible is None:
                first_visible = i
        current = self.inst_list.currentItem()
        if current is None or current.isHidden():
            if first_visible is not None:
                self.inst_list.setCurrentRow(first_visible)
            else:
                self._current_code = None
                self._refresh_selected_panel()

    def _on_institution_selected(
        self,
        current: QListWidgetItem | None,
        previous: QListWidgetItem | None = None,
    ) -> None:
        if previous is not None:
            prev_row = self.inst_list.itemWidget(previous)
            if isinstance(prev_row, InstitutionStatusRowWidget):
                prev_row.set_selected(False)
        if current is not None:
            cur_row = self.inst_list.itemWidget(current)
            if isinstance(cur_row, InstitutionStatusRowWidget):
                cur_row.set_selected(True)
        self._current_code = current.data(Qt.ItemDataRole.UserRole) if current else None
        self._refresh_selected_panel()

    def _set_all_institutions(self, enabled: bool) -> None:
        codes = sorted(self._inst_addresses)
        self._selected_codes = set(codes) if enabled else set()
        self._excluded_addresses = set()
        data_store.set_setting("alwaysRoundUpInstitutions", sorted(self._selected_codes))
        data_store.set_setting("excludeRoundUpAddresses", [])
        for i in range(self.inst_list.count()):
            self._update_inst_item(self.inst_list.item(i))
        self._refresh_inst_summary()
        self._refresh_selected_panel()

    def _set_all_addresses_for_current(self, enabled: bool) -> None:
        code = self._current_code
        if not code:
            return
        excluded = set(data_store.get_setting("excludeRoundUpAddresses") or [])
        selected_codes = set(data_store.get_setting("alwaysRoundUpInstitutions") or [])
        for addr in self._inst_addresses.get(code, []):
            excluded.discard(addr)
        if enabled:
            selected_codes.add(code)
        else:
            selected_codes.discard(code)
        data_store.set_setting("alwaysRoundUpInstitutions", sorted(selected_codes))
        data_store.set_setting("excludeRoundUpAddresses", sorted(excluded))
        self._selected_codes = selected_codes
        self._excluded_addresses = excluded
        for i in range(self.inst_list.count()):
            self._update_inst_item(self.inst_list.item(i))
        self._refresh_inst_summary()
        self._refresh_selected_panel()

    def _populate_dept_percents(self) -> None:
        self._clear_layout_widgets(self.dept_percent_lay)
        by_dept = data_store.get_setting("roundUpPercentByDept") or {}
        products = data_store.get_ref("products") or []
        eligible_dept_keys = {
            p.get("deptKey")
            for p in products
            if p.get("deptKey")
            and p.get("showPcs")
            and (p.get("unit") or "").strip().lower() != "шт"
        }

        shown = 0
        for key, name in data_store.get_department_choices():
            if not key:
                continue
            if key not in eligible_dept_keys:
                continue
            row = QHBoxLayout()
            row.addWidget(QLabel(f"{name}:"))
            spin = QDoubleSpinBox()
            spin.setRange(0, 100)
            spin.setDecimals(1)
            spin.setSuffix(" %")
            spin.setSingleStep(5)
            spin.setValue(float(by_dept.get(key, data_store.get_setting("roundUpInstitutionPercent") or 20)))
            spin.valueChanged.connect(lambda v, k=key: self._on_dept_pct_changed(k, v))
            row.addWidget(spin)
            row.addStretch()
            self.dept_percent_lay.addLayout(row)
            shown += 1

        if shown == 0:
            empty = QLabel("Нет отделов с включённым отображением Шт.")
            empty.setObjectName("hintLabel")
            empty.setWordWrap(True)
            self.dept_percent_lay.addWidget(empty)

    def _set_institution_state(self, code: str, enabled: bool) -> None:
        codes = set(data_store.get_setting("alwaysRoundUpInstitutions") or [])
        if enabled:
            codes.add(code)
        else:
            codes.discard(code)
        data_store.set_setting("alwaysRoundUpInstitutions", sorted(codes))
        self._selected_codes = codes
        for i in range(self.inst_list.count()):
            item = self.inst_list.item(i)
            if item.data(Qt.ItemDataRole.UserRole) == code:
                self._update_inst_item(item)
                break
        self._refresh_inst_summary()
        self._refresh_selected_panel()

    def _on_current_inst_toggle(self, state: int) -> None:
        if self._current_code:
            self._set_institution_state(self._current_code, state == Qt.CheckState.Checked.value)

    def _on_addr_toggle(self, addr: str, state: int) -> None:
        excluded = set(data_store.get_setting("excludeRoundUpAddresses") or [])
        if state != Qt.CheckState.Checked.value:
            excluded.add(addr)
        else:
            excluded.discard(addr)
        data_store.set_setting("excludeRoundUpAddresses", sorted(excluded))
        self._excluded_addresses = excluded
        for i in range(self.inst_list.count()):
            self._update_inst_item(self.inst_list.item(i))
        self._refresh_inst_summary()
        self._refresh_selected_panel()

    def _on_default_pct_changed(self, val: float) -> None:
        data_store.set_setting("roundUpInstitutionPercent", val)

    def _on_dept_pct_changed(self, dept_key: str, val: float) -> None:
        by_dept = dict(data_store.get_setting("roundUpPercentByDept") or {})
        by_dept[dept_key] = val
        data_store.set_setting("roundUpPercentByDept", by_dept)


class InstitutionsDialog(QDialog):
    """Обёртка для совместимости: отдельное окно с тем же встроенным виджетом."""

    def __init__(self, parent: QWidget | None, app_state: dict):
        super().__init__(parent)
        self.setWindowTitle("Округление по учреждениям")
        self.setMinimumSize(820, 640)
        self.resize(920, 720)
        self.setWindowModality(Qt.WindowModality.ApplicationModal)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(12)
        widget = InstitutionsSettingsWidget(app_state, self)
        lay.addWidget(widget)
        btn_close = QPushButton("Закрыть")
        btn_close.setObjectName("btnSecondary")
        btn_close.clicked.connect(self.accept)
        lay.addWidget(btn_close, alignment=Qt.AlignmentFlag.AlignRight)


# ─────────────────────────── Диалог «Настройки Количества» ────────────────────

class QuantitySettingsDialog(QDialog):
    """Окно настроек количества: вкладки отделов, контекст слева, товары и 1 шт справа."""

    def __init__(self, parent: QWidget | None, app_state: dict):
        super().__init__(parent)
        self.setWindowTitle("Настройки Количества")
        self.setMinimumSize(720, 560)
        self.resize(800, 620)
        self.setWindowModality(Qt.WindowModality.ApplicationModal)

        self._app_state = app_state
        self._tab_defs: list[dict] = []
        self._current_tab_key: str | None = None
        self._current_scope_id: str | None = None
        self._current_scope_keys: list[str] = []
        self._scope_panels: dict[str, DeptRoundPanel] = {}
        self._expanded_name: str | None = None
        self._product_panels: dict[str, ProductPcsPanel] = {}
        self._product_row_widgets: dict[str, ProductRowWidget] = {}
        self._pcs_totals_by_product: dict[str, float] = {}

        self._build_ui()
        self._fill_products()

    def _build_ui(self) -> None:
        content = QWidget()
        lay = QVBoxLayout(content)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(16)

        # Заголовок
        lbl_title = QLabel("Настройки Количества")
        lbl_title.setObjectName("sectionTitle")
        lay.addWidget(lbl_title)

        hint = QLabel(
            "Выберите отдел сверху. Если у него есть подотделы, ниже появятся отдельные вкладки подотделов. "
            "Справа задаются проценты округления Шт для школ и садов и настраивается 1 шт по товарам."
        )
        hint.setObjectName("stepLabel")
        hint.setWordWrap(True)
        lay.addWidget(hint)

        inst_card = QFrame()
        inst_card.setObjectName("card")
        inst_lay = QHBoxLayout(inst_card)
        inst_lay.setContentsMargins(16, 14, 16, 14)
        inst_lay.setSpacing(12)

        inst_text = QVBoxLayout()
        inst_text.setSpacing(4)
        inst_title = QLabel("Округление по учреждениям")
        inst_title.setObjectName("sectionTitle")
        inst_text.addWidget(inst_title)
        inst_hint = QLabel(
            "Отдельная настройка для учреждений и адресов сохраняется и работает так же, как раньше."
        )
        inst_hint.setObjectName("hintLabel")
        inst_hint.setWordWrap(True)
        inst_text.addWidget(inst_hint)
        inst_lay.addLayout(inst_text, 1)

        btn_inst = QPushButton("Открыть...")
        btn_inst.setObjectName("btnSecondary")
        btn_inst.clicked.connect(self._open_institutions_dialog)
        inst_lay.addWidget(btn_inst, alignment=Qt.AlignmentFlag.AlignVCenter)
        lay.addWidget(inst_card)

        self.tabs = QTabWidget()
        self.tabs.currentChanged.connect(self._on_tab_changed)
        lay.addWidget(self.tabs)

        self.subtabs = QTabWidget()
        self.subtabs.currentChanged.connect(self._on_subtab_changed)
        self.subtabs.setVisible(False)
        lay.addWidget(self.subtabs)

        self.dept_panel_host = QFrame()
        self.dept_panel_host.setFrameShape(QFrame.Shape.StyledPanel)
        self.dept_panel_lay = QVBoxLayout(self.dept_panel_host)
        self.dept_panel_lay.setContentsMargins(12, 12, 12, 12)
        self.dept_panel_lay.setSpacing(8)
        self.dept_panel_placeholder = QLabel("Выберите вкладку сверху.")
        self.dept_panel_placeholder.setObjectName("hintLabel")
        self.dept_panel_placeholder.setWordWrap(True)
        self.dept_panel_lay.addWidget(self.dept_panel_placeholder)
        lay.addWidget(self.dept_panel_host)

        products_card = QFrame()
        products_card.setObjectName("card")
        products_lay = QVBoxLayout(products_card)
        products_lay.setContentsMargins(16, 16, 16, 16)
        products_lay.setSpacing(12)

        lbl_product_title = QLabel("Настройка 1 шт товара")
        lbl_product_title.setObjectName("sectionTitle")
        products_lay.addWidget(lbl_product_title)

        search_row = QHBoxLayout()
        search_row.addWidget(QLabel("Поиск продукта:"))
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Введите часть названия продукта")
        self.search_edit.textChanged.connect(self._apply_filter)
        search_row.addWidget(self.search_edit)
        products_lay.addLayout(search_row)

        self.lbl_products_context = QLabel("")
        self.lbl_products_context.setObjectName("hintLabel")
        self.lbl_products_context.setWordWrap(True)
        products_lay.addWidget(self.lbl_products_context)

        main_row = QHBoxLayout()
        main_row.setSpacing(12)

        self.products_list = QListWidget()
        self.products_list.setAlternatingRowColors(True)
        self.products_list.setMinimumWidth(260)
        self.products_list.setSpacing(2)
        self.products_list.itemDoubleClicked.connect(self._on_item_double_clicked)
        main_row.addWidget(self.products_list, 1)

        self.panel_stack = QFrame()
        self.panel_stack.setFrameShape(QFrame.Shape.StyledPanel)
        self.panel_lay = QVBoxLayout(self.panel_stack)
        self.panel_lay.setContentsMargins(12, 12, 12, 12)
        self.panel_lay.setSpacing(8)

        self.panel_title = QLabel()
        self.panel_title.setObjectName("sectionTitle")
        self.panel_title.setWordWrap(True)
        self.panel_title.setVisible(False)
        self.panel_lay.addWidget(self.panel_title)

        self.panel_placeholder = QLabel("Дважды нажмите на продукт для настройки 1 шт.")
        self.panel_placeholder.setObjectName("hintLabel")
        self.panel_placeholder.setWordWrap(True)
        self.panel_lay.addWidget(self.panel_placeholder)
        main_row.addWidget(self.panel_stack, 2)

        products_lay.addLayout(main_row)

        lay.addWidget(products_card, 1)

        btn_close = QPushButton("Закрыть")
        btn_close.setObjectName("btnSecondary")
        btn_close.clicked.connect(self.accept)
        lay.addWidget(btn_close, alignment=Qt.AlignmentFlag.AlignRight)

        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setWidget(content)
        main_lay = QVBoxLayout(self)
        main_lay.setContentsMargins(0, 0, 0, 0)
        main_lay.addWidget(scroll)

    def _clear_panel_widgets(self, layout: QVBoxLayout, keep_first: int = 0) -> None:
        while layout.count() > keep_first:
            item = layout.takeAt(keep_first)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

    def _eligible_products(self) -> list[dict]:
        products = data_store.get_ref("products") or []
        return sorted(
            [
                p for p in products
                if p.get("deptKey") and p.get("name")
                and (p.get("unit") or "").strip().lower() != "шт"
            ],
            key=lambda p: (p.get("name") or "").lower()
        )

    def _build_tab_defs(self, eligible_products: list[dict]) -> list[dict]:
        counts_by_key: dict[str, int] = {}
        for prod in eligible_products:
            dept_key = prod.get("deptKey")
            if dept_key:
                counts_by_key[dept_key] = counts_by_key.get(dept_key, 0) + 1

        result: list[dict] = []
        for dept in data_store.get_ref("departments") or []:
            dept_key = dept.get("key") or ""
            if not dept_key:
                continue
            dept_name = dept.get("name") or dept_key
            direct_count = counts_by_key.get(dept_key, 0)
            sub_tabs: list[dict] = []
            total_count = direct_count
            for sub in dept.get("subdepts", []):
                sub_key = sub.get("key") or ""
                sub_name = sub.get("name") or sub_key
                sub_count = counts_by_key.get(sub_key, 0)
                if not sub_key or sub_count <= 0:
                    continue
                total_count += sub_count
                sub_tabs.append(
                    {
                        "scope_id": f"sub:{sub_key}",
                        "tab_text": f"{sub_name} ({sub_count})",
                        "scope_label": sub_name,
                        "dept_keys": [sub_key],
                        "count": sub_count,
                    }
                )
            if total_count <= 0:
                continue
            result.append(
                {
                    "scope_id": f"depttab:{dept_key}",
                    "tab_text": f"{dept_name} ({total_count})",
                    "dept_key": dept_key,
                    "scope_label": dept_name,
                    "count": total_count,
                    "default_scope": {
                        "scope_id": f"dept:{dept_key}",
                        "tab_text": f"Отдел ({direct_count})",
                        "scope_label": f"{dept_name} - товары отдела",
                        "dept_keys": [dept_key],
                        "count": direct_count,
                    },
                    "sub_tabs": sub_tabs,
                }
            )
        return result

    def _fill_products(self) -> None:
        """Заполняет товары и вкладки отделов."""
        self.products_list.clear()
        self.tabs.clear()
        self.subtabs.clear()
        self.subtabs.setVisible(False)
        self._tab_defs.clear()
        self._current_tab_key = None
        self._current_scope_id = None
        self._current_scope_keys = []
        self._clear_panel_widgets(self.dept_panel_lay, keep_first=1)
        self._scope_panels.clear()
        self._clear_panel_widgets(self.panel_lay, keep_first=2)
        self._product_panels.clear()
        self._product_row_widgets.clear()
        eligible = self._eligible_products()
        for prod in eligible:
            name = prod["name"]
            item = QListWidgetItem()
            item.setData(Qt.ItemDataRole.UserRole, name)
            row = ProductRowWidget(name, self.products_list)
            self._product_row_widgets[name] = row
            self.products_list.addItem(item)
            self.products_list.setItemWidget(item, row)
            item.setSizeHint(QSize(self.products_list.width() or 260, ProductRowWidget.ROW_HEIGHT))
        self._tab_defs = self._build_tab_defs(eligible)
        self._populate_tabs()
        self._refresh_all_pcs_totals()
        self._apply_filter(self.search_edit.text())

    def _populate_tabs(self) -> None:
        for tab_def in self._tab_defs:
            self.tabs.addTab(QWidget(), tab_def["tab_text"])
        if not self._tab_defs:
            self._show_dept_placeholder()
            self.lbl_products_context.setText("Нет отделов или подотделов с товарами для настройки Шт.")
            return
        selected_index = 0
        if self._current_tab_key:
            for i, tab_def in enumerate(self._tab_defs):
                if tab_def["dept_key"] == self._current_tab_key:
                    selected_index = i
                    break
        self.tabs.setCurrentIndex(selected_index)
        self._on_tab_changed(selected_index)

    def _show_dept_placeholder(self) -> None:
        for panel in self._scope_panels.values():
            panel.setVisible(False)
        self.dept_panel_placeholder.setVisible(True)

    def _show_scope_panel(self, scope_def: dict) -> None:
        scope_id = scope_def["scope_id"]
        if scope_id not in self._scope_panels:
            panel = DeptRoundPanel(
                scope_id=scope_id,
                scope_label=scope_def["scope_label"],
                dept_keys=scope_def["dept_keys"],
                on_changed=self._on_scope_rounding_changed,
                parent=self,
            )
            self._scope_panels[scope_id] = panel
            self.dept_panel_lay.addWidget(panel)
        self._scope_panels[scope_id]._load()
        for key, panel in self._scope_panels.items():
            panel.setVisible(key == scope_id)
        self.dept_panel_placeholder.setVisible(False)

    def _populate_subtabs(self, tab_def: dict) -> None:
        self.subtabs.blockSignals(True)
        self.subtabs.clear()
        self.subtabs.addTab(QWidget(), tab_def["default_scope"]["tab_text"])
        for scope_def in tab_def.get("sub_tabs", []):
            self.subtabs.addTab(QWidget(), scope_def["tab_text"])
        has_subtabs = bool(tab_def.get("sub_tabs"))
        self.subtabs.setVisible(has_subtabs)
        self.subtabs.setCurrentIndex(0)
        self.subtabs.blockSignals(False)

    def _apply_scope(self, scope_def: dict | None) -> None:
        self._current_scope_id = scope_def["scope_id"] if scope_def else None
        self._current_scope_keys = list(scope_def.get("dept_keys", [])) if scope_def else []
        if scope_def:
            self._show_scope_panel(scope_def)
            self.lbl_products_context.setText(f"Товары блока: {scope_def['scope_label']}.")
        else:
            self._show_dept_placeholder()
            self.lbl_products_context.setText("Сначала выберите вкладку сверху.")
        self._apply_filter(self.search_edit.text())

    def _on_tab_changed(self, index: int) -> None:
        if index < 0 or index >= len(self._tab_defs):
            self._current_tab_key = None
            self._current_scope_id = None
            self._current_scope_keys = []
            self.subtabs.clear()
            self.subtabs.setVisible(False)
            self._show_dept_placeholder()
            self.lbl_products_context.setText("Нет отделов для настройки.")
            self._apply_filter(self.search_edit.text())
            return
        tab_def = self._tab_defs[index]
        self._current_tab_key = tab_def["dept_key"]
        self._populate_subtabs(tab_def)
        self._apply_scope(tab_def["default_scope"])

    def _on_subtab_changed(self, index: int) -> None:
        if self.tabs.currentIndex() < 0 or self.tabs.currentIndex() >= len(self._tab_defs):
            return
        tab_def = self._tab_defs[self.tabs.currentIndex()]
        if index <= 0:
            self._apply_scope(tab_def["default_scope"])
            return
        sub_tabs = tab_def.get("sub_tabs", [])
        sub_index = index - 1
        if 0 <= sub_index < len(sub_tabs):
            self._apply_scope(sub_tabs[sub_index])

    def _on_scope_rounding_changed(self, scope_id: str) -> None:
        if scope_id in self._scope_panels:
            self._scope_panels[scope_id]._load()
        if self._expanded_name:
            products = data_store.get_ref("products") or []
            prod = next((p for p in products if p.get("name") == self._expanded_name), None)
            if (
                prod
                and prod.get("deptKey") in self._current_scope_keys
                and self._expanded_name in self._product_panels
            ):
                self._product_panels[self._expanded_name]._load()
        self._refresh_all_pcs_totals()

    def _refresh_all_pcs_totals(self) -> None:
        self._pcs_totals_by_product = _calc_product_pcs_totals(self._app_state)
        products_ref = data_store.get_ref("products") or []
        by_name = {p.get("name"): p for p in products_ref if p.get("name")}
        for name, row in self._product_row_widgets.items():
            prod = by_name.get(name) or {}
            total_text = ""
            if prod.get("showPcs"):
                total_text = f"Всего по текущим маршрутам: {_format_pcs_total(self._pcs_totals_by_product.get(name, 0.0))}"
            row.set_total_pcs_text(total_text)
            mode_text = ""
            if excel_generator.get_dept_special_mode(prod.get("deptKey")) == "polufabricates":
                mode_text = "Режим: Полуфабрикаты (без округления)"
            row.set_mode_text(mode_text)
        for name, panel in self._product_panels.items():
            prod = by_name.get(name) or {}
            total_text = ""
            if prod.get("showPcs"):
                total_text = f"Всего по текущим маршрутам: {_format_pcs_total(self._pcs_totals_by_product.get(name, 0.0))}"
            panel.set_total_pcs_text(total_text)

    def _on_product_settings_changed(self, product_name: str) -> None:
        del product_name
        if self._current_scope_id and self._current_scope_id in self._scope_panels:
            self._scope_panels[self._current_scope_id]._load()
        self._refresh_all_pcs_totals()

    def _on_item_double_clicked(self, list_item: QListWidgetItem) -> None:
        """По двойному клику открывает/закрывает панель настроек в Шт для продукта."""
        name = list_item.data(Qt.ItemDataRole.UserRole)
        if name:
            self._open_product_panel(name)

    def _open_product_panel(self, product_name: str) -> None:
        """Переключает видимость панели настроек продукта (открыть или закрыть при повторе)."""
        if self._expanded_name == product_name:
            self._expanded_name = None
            self._show_panel_placeholder()
            return
        self._expanded_name = product_name
        if product_name not in self._product_panels:
            pan = ProductPcsPanel(product_name, on_changed=self._on_product_settings_changed, parent=self)
            self._product_panels[product_name] = pan
            self.panel_lay.addWidget(pan)
            self._refresh_all_pcs_totals()
        self._product_panels[product_name]._load()
        for name, pan in self._product_panels.items():
            pan.setVisible(name == product_name)
        self.panel_placeholder.hide()
        self.panel_title.setText(product_name)
        self.panel_title.setVisible(True)

    def _show_panel_placeholder(self) -> None:
        for pan in self._product_panels.values():
            pan.setVisible(False)
        self.panel_title.setVisible(False)
        self.panel_placeholder.setVisible(True)

    def _apply_filter(self, text: str) -> None:
        """Фильтрация списка продуктов по текущему блоку и подстроке."""
        pattern = (text or "").strip().lower()
        products = data_store.get_ref("products") or []
        dept_by_name = {p.get("name"): p.get("deptKey") for p in products if p.get("name")}
        for i in range(self.products_list.count()):
            item = self.products_list.item(i)
            name = (item.data(Qt.ItemDataRole.UserRole) or "").lower()
            item_name = item.data(Qt.ItemDataRole.UserRole) or ""
            visible = (
                (not pattern or pattern in name)
                and (not self._current_scope_keys or dept_by_name.get(item_name) in self._current_scope_keys)
            )
            item.setHidden(not visible)
        # Если текущий раскрытый продукт скрыт фильтром — свернуть панель
        if self._expanded_name:
            any_visible = False
            for i in range(self.products_list.count()):
                item = self.products_list.item(i)
                if not item.isHidden() and (
                    item.data(Qt.ItemDataRole.UserRole) == self._expanded_name
                ):
                    any_visible = True
                    break
            if not any_visible:
                self._expanded_name = None
                self._show_panel_placeholder()

    def _open_institutions_dialog(self) -> None:
        dlg = InstitutionsDialog(self, self._app_state)
        dlg.exec()


def open_quantity_settings_dialog(parent: QWidget, app_state: dict) -> None:
    """Открывает диалог «Настройки Количества»."""
    dlg = QuantitySettingsDialog(parent, app_state)
    dlg.exec()
