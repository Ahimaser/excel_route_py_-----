"""
quantity_settings_dialog.py — Настройки Количества.

Оглавление секций:
  1. Утилиты (_extract_institution_list, ProductPcsPanel, ProductCardWidget)
  2. QuantitySettingsDialog: UI (вкладки отделов, карточки продуктов)
  3. Панель настроек Шт (ProductPcsPanel), округление по учреждениям
  4. Сохранение и обновление данных

Окно: список продуктов (привязанных к отделам). По двойному клику — настройка Шт справа.
Ниже — блок «Округление по учреждениям».
"""
from __future__ import annotations

import copy
from typing import Iterable

from PyQt6.QtCore import Qt, QSize, pyqtSignal
from PyQt6.QtGui import QKeySequence, QShortcut
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
    QGridLayout,
    QComboBox,
)

from core import data_store, excel_generator
from ui.widgets import ToggleSwitch, hint_icon_button


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
    """Шт — всегда целое число."""
    if total is None:
        return ""
    return f"{int(round(float(total)))} шт"


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
        self.lbl_unit_pcs = QLabel("")
        self.lbl_unit_pcs.setObjectName("unitLabel")
        row2.addWidget(self.lbl_unit_pcs)
        row2.addStretch()
        lay.addLayout(row2)

        # От кол-ва в 1 шт (минимальное для округления)
        row_min = QHBoxLayout()
        row_min.addWidget(QLabel("От кол-ва в 1 шт:"))
        self.spin_min_qty = QDoubleSpinBox()
        self.spin_min_qty.setRange(0, 99999.0)
        self.spin_min_qty.setDecimals(3)
        self.spin_min_qty.setSingleStep(0.1)
        self.spin_min_qty.setToolTip("Ниже этого количества — 0 шт. От этого и выше — расчёт по «Кол-во в 1 шт» и округлению.")
        self.spin_min_qty.valueChanged.connect(self._on_min_qty_changed)
        row_min.addWidget(self.spin_min_qty)
        self.lbl_unit_min = QLabel("")
        self.lbl_unit_min.setObjectName("unitLabel")
        row_min.addWidget(self.lbl_unit_min)
        row_min.addStretch()
        lay.addLayout(row_min)

        # В Грязные (только подотдел Чищенка: конвертация 1,25, без этикеток, колонка «Грязные»)
        self.row_dirty_widget = QWidget()
        row_dirty = QHBoxLayout(self.row_dirty_widget)
        row_dirty.setContentsMargins(0, 0, 0, 0)
        self.chk_show_in_dirty = ToggleSwitch()
        self.chk_show_in_dirty.setToolTip(
            "Только для подотдела «Чищенка». Отображать в «Грязные» (×1,25). "
            "Этикетки на продукт не печатаются. В таблицах добавляется колонка «Грязные»."
        )
        self.chk_show_in_dirty.stateChanged.connect(self._on_show_in_dirty_changed)
        row_dirty.addWidget(QLabel("В Грязные:"))
        row_dirty.addWidget(self.chk_show_in_dirty)
        row_dirty.addStretch()
        lay.addWidget(self.row_dirty_widget)
        self.row_dirty_widget.setVisible(False)  # по умолчанию скрыто, показывается только для Чищенка

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
            self.spin_min_qty.setEnabled(False)
            self.row_dirty_widget.setVisible(False)
            self.lbl_unit_pcs.setText("")
            self.lbl_unit_min.setText("")
            self._loading = False
            return
        unit = (prod.get("unit") or "").strip()
        unit_lower = unit.lower()
        if unit_lower == "шт":
            self.chk_show.setEnabled(False)
            self.spin_pcs.setEnabled(False)
            self.spin_min_qty.setEnabled(False)
            self.row_dirty_widget.setVisible(False)
            self.lbl_unit_pcs.setText("")
            self.lbl_unit_min.setText("")
            self._loading = False
            return
        self.lbl_unit_pcs.setText(unit)
        self.lbl_unit_min.setText(unit)
        show_pcs = prod.get("showPcs", False)
        self.chk_show.setChecked(show_pcs)
        pcs_per_unit = float(prod.get("pcsPerUnit", 1.0) or 1.0)
        self.spin_pcs.setValue(pcs_per_unit)
        self.spin_pcs.setEnabled(show_pcs)
        self.spin_min_qty.setValue(prod.get("minQtyForPcs", 0) or 0)
        self.spin_min_qty.setEnabled(show_pcs)
        in_chistchenka = data_store.is_subdept_chistchenka(prod.get("deptKey"))
        self.row_dirty_widget.setVisible(in_chistchenka)
        self.chk_show_in_dirty.setChecked(bool(in_chistchenka and prod.get("showInDirty", False)))
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
        self.spin_min_qty.setEnabled(show)
        data_store.update_product(self._product_name, showPcs=show)
        self._notify_changed()

    def _on_min_qty_changed(self, val: float) -> None:
        if self._loading:
            return
        data_store.update_product(self._product_name, minQtyForPcs=val if val > 0 else None)
        self._notify_changed()

    def _on_show_in_dirty_changed(self, state: int) -> None:
        if self._loading:
            return
        products = data_store.get_ref("products") or []
        prod = next((p for p in products if p.get("name") == self._product_name), None)
        if not prod or not data_store.is_subdept_chistchenka(prod.get("deptKey")):
            return
        enabled = state == Qt.CheckState.Checked.value
        data_store.update_product(
            self._product_name,
            showInDirty=enabled,
            quantityMultiplier=1.25 if enabled else None,
        )
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


# ─────────────────────────── Окно настройки продукта (вариант F) ─────────────

class ProductPcsSettingsDialog(QDialog):
    """Отдельное модальное окно настройки Шт для одного продукта."""

    open_next_requested = pyqtSignal(str)

    def __init__(
        self,
        product_name: str,
        product_list: list[str],
        app_state: dict,
        on_changed_callback=None,
        parent: QWidget | None = None,
    ):
        super().__init__(parent)
        self._product_name = product_name
        self._product_list = product_list
        self._on_changed_cb = on_changed_callback
        self.setWindowTitle(f"Настройки: {product_name}")
        self.setMinimumSize(380, 320)
        self.resize(420, 380)
        self.setWindowModality(Qt.WindowModality.ApplicationModal)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 16, 20, 16)
        lay.setSpacing(12)

        self._panel = ProductPcsPanel(product_name, on_changed=self._on_panel_changed, parent=self)
        lay.addWidget(self._panel)

        pcs_totals = _calc_product_pcs_totals(app_state)
        products_ref = data_store.get_ref("products") or []
        by_name = {p.get("name"): p for p in products_ref if p.get("name")}
        prod = by_name.get(product_name) or {}
        if prod.get("showPcs"):
            total_text = f"Всего по маршрутам: {_format_pcs_total(pcs_totals.get(product_name, 0.0))}"
            self._panel.set_total_pcs_text(total_text)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self._btn_next = QPushButton("Настроить следующий")
        self._btn_next.setObjectName("btnSecondary")
        self._btn_next.clicked.connect(self._on_next_clicked)
        self._update_next_button()
        btn_row.addWidget(self._btn_next)
        btn_save = QPushButton("Сохранить")
        btn_save.setObjectName("btnPrimary")
        btn_save.setDefault(True)
        btn_save.setAutoDefault(True)
        btn_save.clicked.connect(self.accept)
        btn_row.addWidget(btn_save)
        lay.addLayout(btn_row)

        QShortcut(QKeySequence(Qt.Key.Key_Escape), self, self.accept)
        QShortcut(QKeySequence(Qt.Key.Key_Return), self, self.accept)

    def _update_next_button(self) -> None:
        idx = self._product_list.index(self._product_name) if self._product_name in self._product_list else -1
        has_next = 0 <= idx < len(self._product_list) - 1
        self._btn_next.setVisible(has_next)

    def _on_panel_changed(self, product_name: str) -> None:
        if callable(self._on_changed_cb):
            self._on_changed_cb(product_name)

    def _on_next_clicked(self) -> None:
        idx = self._product_list.index(self._product_name) if self._product_name in self._product_list else -1
        if 0 <= idx < len(self._product_list) - 1:
            next_name = self._product_list[idx + 1]
            self.open_next_requested.emit(next_name)
            # Слот родителя закроет диалог и откроет следующий


# ─────────────────────────── Строка списка: название продукта ────────────────

class ProductRowWidget(QWidget):
    """Одна строка: название продукта (настройки открываются по двойному клику)."""

    ROW_HEIGHT = 70

    def __init__(self, product_name: str, parent=None, prod_map: dict | None = None):
        super().__init__(parent)
        self.product_name = product_name
        self.setMinimumHeight(self.ROW_HEIGHT)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(12, 12, 12, 12)
        text_col = QVBoxLayout()
        text_col.setSpacing(2)

        display_name = data_store.format_product_display_name(product_name, prod_map or {}) if prod_map else product_name
        self.lbl_name = QLabel(display_name)
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


class ProductCardWidget(QFrame):
    """Карточка продукта: название, бейдж единицы, превью настроек. Клик — выбор."""

    clicked = pyqtSignal(str)

    def __init__(self, product_name: str, parent=None, prod_map: dict | None = None):
        super().__init__(parent)
        self.product_name = product_name
        self._selected = False
        self._disabled_for_pcs = False  # True для продуктов с ед. изм. «шт»
        self.setObjectName("productCard")
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setMinimumHeight(120)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(14, 12, 14, 12)
        lay.setSpacing(6)

        top_row = QHBoxLayout()
        display_name = data_store.format_product_display_name(product_name, prod_map or {}) if prod_map else product_name
        self.lbl_name = QLabel(display_name)
        self.lbl_name.setObjectName("cardTitle")
        self.lbl_name.setWordWrap(True)
        self.lbl_name.setAlignment(Qt.AlignmentFlag.AlignTop)
        font = QFont()
        font.setPointSize(14)
        font.setWeight(500)
        self.lbl_name.setFont(font)
        top_row.addWidget(self.lbl_name, 1)
        self.lbl_badge = QLabel("")
        self.lbl_badge.setObjectName("badge")
        self.lbl_badge.setVisible(False)
        top_row.addWidget(self.lbl_badge, 0, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignRight)
        lay.addLayout(top_row)

        self.lbl_preview = QLabel("")
        self.lbl_preview.setObjectName("cardPreview")
        self.lbl_preview.setWordWrap(True)
        lay.addWidget(self.lbl_preview)

    def set_preview_text(self, text: str) -> None:
        self.lbl_preview.setText(text)
        self.lbl_preview.setVisible(bool(text))

    def set_unit_badge(self, unit: str) -> None:
        if unit:
            self.lbl_badge.setText(unit)
            self.lbl_badge.setVisible(True)
        else:
            self.lbl_badge.setVisible(False)

    def set_tooltip_text(self, text: str) -> None:
        self.setToolTip(text)

    def set_selected(self, selected: bool) -> None:
        self._selected = selected
        self.setObjectName("productCardSelected" if selected else "productCard")
        self.style().unpolish(self)
        self.style().polish(self)

    def set_disabled_for_pcs(self, disabled: bool) -> None:
        """Отключает карточку для продуктов с ед. изм. «шт» — полупрозрачная, без клика."""
        self._disabled_for_pcs = disabled
        self.setEnabled(not disabled)
        self.setCursor(Qt.CursorShape.ForbiddenCursor if disabled else Qt.CursorShape.PointingHandCursor)
        self.setGraphicsEffect(None)
        if disabled:
            from PyQt6.QtWidgets import QGraphicsOpacityEffect
            eff = QGraphicsOpacityEffect(self)
            eff.setOpacity(0.45)
            self.setGraphicsEffect(eff)

    def mousePressEvent(self, event) -> None:
        super().mousePressEvent(event)
        if event.button() == Qt.MouseButton.LeftButton and not self._disabled_for_pcs:
            self.clicked.emit(self.product_name)


class DeptRoundPanel(QFrame):
    """Панель настройки округления хвостика для Школ и Садов: направление (>/<), значение в %."""

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
        lay.setContentsMargins(12, 10, 12, 10)
        lay.setSpacing(12)

        self.lbl_title = QLabel(scope_label)
        self.lbl_title.setObjectName("cardTitle")
        self.lbl_title.setWordWrap(False)
        lay.addWidget(self.lbl_title)

        row_shk = QHBoxLayout()
        row_shk.setSpacing(12)
        row_shk.addWidget(QLabel("Округление хвостика для Школ от %:"))
        self.combo_dir_shk = QComboBox()
        self.combo_dir_shk.addItems(["> (вверх)", "< (вниз)"])
        self.combo_dir_shk.setMinimumWidth(140)
        self.combo_dir_shk.setMinimumHeight(28)
        self.combo_dir_shk.currentIndexChanged.connect(self._on_shk_ui_changed)
        row_shk.addWidget(self.combo_dir_shk)
        self.spin_shk = QDoubleSpinBox()
        self.spin_shk.setRange(0, 100.0)
        self.spin_shk.setDecimals(1)
        self.spin_shk.setSuffix(" %")
        self.spin_shk.setSingleStep(5.0)
        self.spin_shk.setMinimumWidth(90)
        self.spin_shk.setMinimumHeight(28)
        self.spin_shk.valueChanged.connect(self._on_shk_changed)
        row_shk.addWidget(self.spin_shk)
        row_shk.addStretch()
        lay.addLayout(row_shk)

        row_sd = QHBoxLayout()
        row_sd.setSpacing(12)
        row_sd.addWidget(QLabel("Округление хвостика для Садов от %:"))
        self.combo_dir_sd = QComboBox()
        self.combo_dir_sd.addItems(["> (вверх)", "< (вниз)"])
        self.combo_dir_sd.setMinimumWidth(140)
        self.combo_dir_sd.setMinimumHeight(28)
        self.combo_dir_sd.currentIndexChanged.connect(self._on_sd_ui_changed)
        row_sd.addWidget(self.combo_dir_sd)
        self.spin_sd = QDoubleSpinBox()
        self.spin_sd.setRange(0, 100.0)
        self.spin_sd.setDecimals(1)
        self.spin_sd.setSuffix(" %")
        self.spin_sd.setSingleStep(5.0)
        self.spin_sd.setMinimumWidth(90)
        self.spin_sd.setMinimumHeight(28)
        self.spin_sd.valueChanged.connect(self._on_sd_changed)
        row_sd.addWidget(self.spin_sd)
        row_sd.addStretch()
        lay.addLayout(row_sd)

        self.lbl_mixed = QLabel("")
        self.lbl_mixed.setObjectName("stepLabel")
        self.lbl_mixed.setWordWrap(True)
        self.lbl_mixed.setVisible(False)
        lay.addWidget(self.lbl_mixed)
        self._load()

    def _dept_products(self) -> list[dict]:
        products = data_store.get_ref("products") or []
        return [
            p for p in products
            if p.get("deptKey") in self._dept_keys
            and p.get("name")
            and (p.get("unit") or "").strip().lower() != "шт"
        ]

    def _product_tail_value(self, prod: dict, route_cat: str) -> float | None:
        pcs_per_unit = float(prod.get("pcsPerUnit", 1.0) or 1.0)
        if route_cat == "СД":
            tail_value = prod.get("roundTailFromСД")
            round_up = prod.get("roundUpСД", True)
        else:
            tail_value = prod.get("roundTailFromШК")
            round_up = prod.get("roundUpШК", True)
        if tail_value is not None:
            return float(tail_value)
        return 0.0 if round_up else pcs_per_unit

    def _product_direction(self, prod: dict, route_cat: str) -> bool:
        """True = вверх (>), False = вниз (<)."""
        if route_cat == "СД":
            return prod.get("roundUpСД", True)
        return prod.get("roundUpШК", True)

    def _load(self) -> None:
        self._loading = True
        products = self._dept_products()
        self.lbl_title.setText(self._scope_label)

        shk_tails = [self._product_tail_value(p, "ШК") for p in products]
        sd_tails = [self._product_tail_value(p, "СД") for p in products]
        shk_dirs = [self._product_direction(p, "ШК") for p in products]
        sd_dirs = [self._product_direction(p, "СД") for p in products]

        shk_up = all(shk_dirs) if shk_dirs else True
        sd_up = all(sd_dirs) if sd_dirs else True
        self.combo_dir_shk.setCurrentIndex(0 if shk_up else 1)
        self.combo_dir_sd.setCurrentIndex(0 if sd_up else 1)

        first = products[0] if products else None
        if first:
            pcu = float(first.get("pcsPerUnit", 1.0) or 1.0)
            tv_shk = self._product_tail_value(first, "ШК")
            tv_sd = self._product_tail_value(first, "СД")
            self.spin_shk.setValue(
                round(_tail_value_to_percent(tv_shk or 0, pcu), 1) if shk_up else 0.0
            )
            self.spin_sd.setValue(
                round(_tail_value_to_percent(tv_sd or 0, pcu), 1) if sd_up else 0.0
            )
        else:
            self.spin_shk.setValue(0.0)
            self.spin_sd.setValue(0.0)

        self._set_spin_enabled("ШК", shk_up)
        self._set_spin_enabled("СД", sd_up)

        mixed_parts = []
        shk_vals = sorted({self._product_tail_value(p, "ШК") for p in products}) or [0.0]
        sd_vals = sorted({self._product_tail_value(p, "СД") for p in products}) or [0.0]
        if len(shk_vals) > 1:
            mixed_parts.append("Школы")
        if len(sd_vals) > 1:
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

    def _set_spin_enabled(self, route_cat: str, enabled: bool) -> None:
        """Включает/выключает только спинбокс. Комбобокс %/ед. всегда активен для выбора."""
        if route_cat == "ШК":
            self.spin_shk.setEnabled(enabled)
        else:
            self.spin_sd.setEnabled(enabled)

    def _apply_rounding(self, route_cat: str, direction_up: bool, value: float) -> None:
        """Применяет округление. value всегда в % от 1 шт."""
        ref = data_store.get_ref("products")
        if ref is None:
            return
        products = copy.deepcopy(ref)
        tail_field = "roundTailFromСД" if route_cat == "СД" else "roundTailFromШК"
        up_field = "roundUpСД" if route_cat == "СД" else "roundUpШК"
        changed = False
        for prod in products:
            if prod.get("deptKey") not in self._dept_keys:
                continue
            if not prod.get("name"):
                continue
            if (prod.get("unit") or "").strip().lower() == "шт":
                continue
            pcs_per_unit = float(prod.get("pcsPerUnit", 1.0) or 1.0)
            if direction_up:
                tail_val = _percent_to_tail_value(value, pcs_per_unit)
                prod[tail_field] = max(0.0, float(tail_val))
                prod[up_field] = True
            else:
                prod[tail_field] = None
                prod[up_field] = False
            changed = True
        if not changed:
            return
        data_store.set_key("products", products)
        self._load()
        if callable(self._on_changed_cb):
            self._on_changed_cb(self._scope_id)

    def _on_shk_ui_changed(self, idx: int) -> None:
        if self._loading:
            return
        up = idx == 0
        self._set_spin_enabled("ШК", up)
        if up:
            self._apply_rounding("ШК", True, self.spin_shk.value())
        else:
            self._apply_rounding("ШК", False, 0)

    def _on_sd_ui_changed(self, idx: int) -> None:
        if self._loading:
            return
        up = idx == 0
        self._set_spin_enabled("СД", up)
        if up:
            self._apply_rounding("СД", True, self.spin_sd.value())
        else:
            self._apply_rounding("СД", False, 0)

    def _on_shk_changed(self, val: float) -> None:
        if self._loading:
            return
        if self.combo_dir_shk.currentIndex() != 0:
            return
        self._apply_rounding("ШК", True, val)

    def _on_sd_changed(self, val: float) -> None:
        if self._loading:
            return
        if self.combo_dir_sd.currentIndex() != 0:
            return
        self._apply_rounding("СД", True, val)


class InstitutionStatusRowWidget(QFrame):
    """Строка учреждения с заметным бейджем статуса."""

    ROW_HEIGHT = 72

    def __init__(self, code: str, parent=None):
        super().__init__(parent)
        self.code = code
        self._selected = False
        self.setObjectName("instCard")
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
        self.lbl_badge.setObjectName(
            "badgeGreen" if status == "Включено" else
            "badgeOrange" if status == "Частично" else "badgeGray"
        )
        self.style().unpolish(self.lbl_badge)
        self.style().polish(self.lbl_badge)
        self._apply_selected_style()

    def set_selected(self, selected: bool) -> None:
        self._selected = selected
        self._apply_selected_style()

    def _apply_selected_style(self) -> None:
        self.setObjectName("instCard")
        self.setProperty("selected", self._selected)
        self.style().unpolish(self)
        self.style().polish(self)


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
            "Выберите учреждение слева — справа отобразятся его адреса. "
            "При необходимости исключите отдельные адреса из округления."
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
        hint_dept = QLabel("Для каждого отдела можно задать свой порог округления. Если не задан — используется общий процент.")
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
            self.lbl_current_inst.setObjectName("sectionTitle")
            self.style().unpolish(self.lbl_current_inst)
            self.style().polish(self.lbl_current_inst)
            self.lbl_current_status.setText("")
            self.lbl_current_status.setObjectName("hintLabel")
            self.style().unpolish(self.lbl_current_status)
            self.style().polish(self.lbl_current_status)
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
        self.lbl_current_inst.setObjectName(
            "instTitleGreen" if status == "Включено" else
            "instTitleOrange" if status == "Частично" else "instTitleGray"
        )
        self.style().unpolish(self.lbl_current_inst)
        self.style().polish(self.lbl_current_inst)
        self.lbl_current_status.setText(self._status_badge_text(status))
        self.lbl_current_status.setObjectName(
            "badgeGreen" if status == "Включено" else
            "badgeOrange" if status == "Частично" else "badgeGray"
        )
        self.style().unpolish(self.lbl_current_status)
        self.style().polish(self.lbl_current_status)
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
            # Подотдел «Полуфабрикаты» — округление не применяется, не показывать
            if excel_generator.get_dept_special_mode(key) == "polufabricates":
                continue
            if "полуфаб" in (name or "").lower():
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
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_cancel = QPushButton("Отмена")
        btn_cancel.setObjectName("btnSecondary")
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_cancel)
        btn_save = QPushButton("Сохранить")
        btn_save.setObjectName("btnPrimary")
        btn_save.setDefault(True)
        btn_save.setAutoDefault(True)
        btn_save.clicked.connect(self.accept)
        btn_row.addWidget(btn_save)
        lay.addLayout(btn_row)
        QShortcut(QKeySequence(Qt.Key.Key_Return), self, self.accept)


# ─────────────────────────── Диалог «Настройки Количества» ────────────────────

class QuantitySettingsDialog(QDialog):
    """Окно настроек: отделы сверху, подотделы при выборе, карточки продуктов."""

    def __init__(self, parent: QWidget | None, app_state: dict):
        super().__init__(parent)
        self.setWindowTitle("Настройки Количества")
        self.setObjectName("quantitySettingsDialog")
        self.setMinimumSize(700, 560)
        self.resize(900, 680)
        self.setWindowModality(Qt.WindowModality.ApplicationModal)

        self._app_state = app_state
        self._tab_defs: list[dict] = []
        self._current_tab_key: str | None = None
        self._current_scope_id: str | None = None
        self._current_scope_keys: list[str] = []
        self._scope_panels: dict[str, DeptRoundPanel] = {}
        self._product_cards: dict[str, ProductCardWidget] = {}
        self._pcs_totals_by_product: dict[str, float] = {}
        self._dept_buttons: list[QPushButton] = []
        self._subdept_buttons: list[QPushButton] = []
        self._subdept_scopes: list[dict] = []

        self._build_ui()
        self._fill_data()
        QShortcut(QKeySequence(Qt.Key.Key_Escape), self, self.reject)

    def _build_ui(self) -> None:
        content = QWidget()
        lay = QVBoxLayout(content)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(16)

        title_row = QHBoxLayout()
        lbl_title = QLabel("Настройки Количества")
        lbl_title.setObjectName("sectionTitle")
        title_row.addWidget(lbl_title)
        title_row.addWidget(hint_icon_button(
            self,
            "Настройка отображения в штуках и округления по отделам и учреждениям.",
            "Инструкция — Настройки Количества\n\n"
            "1. Выберите отдел — отобразятся подотделы и карточки продуктов.\n"
            "2. Клик по карточке продукта — настройка Шт (кол-во в 1 шт, хвостики ШК/СД).\n"
            "3. Панель округления — общие настройки хвостика для отдела/подотдела.\n"
            "4. «Округление по учреждениям» — отдельная настройка по учреждениям и адресам.",
            "Инструкция",
        ))
        title_row.addStretch()
        lay.addLayout(title_row)

        hint = QLabel(
            "Выберите отдел — подотделы отобразятся под кнопками отделов. "
            "Товары отдела — без подотделов; товары подотдела — только при его выборе. "
            "Клик по карточке продукта — настройка количества в 1 шт."
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
        inst_hint = QLabel("Отдельная настройка округления для учреждений и адресов.")
        inst_hint.setObjectName("hintLabel")
        inst_hint.setWordWrap(True)
        inst_text.addWidget(inst_hint)
        inst_lay.addLayout(inst_text, 1)
        btn_inst = QPushButton("Открыть...")
        btn_inst.setObjectName("btnSecondary")
        btn_inst.clicked.connect(self._open_institutions_dialog)
        inst_lay.addWidget(btn_inst, alignment=Qt.AlignmentFlag.AlignVCenter)
        lay.addWidget(inst_card)

        dept_row = QHBoxLayout()
        dept_row.setSpacing(8)
        dept_tabs_frame = QFrame()
        dept_tabs_frame.setObjectName("deptTabsBar")
        dept_btns_lay = QHBoxLayout(dept_tabs_frame)
        dept_btns_lay.setContentsMargins(8, 6, 8, 6)
        dept_btns_lay.setSpacing(6)
        self.dept_buttons_widget = dept_tabs_frame
        dept_row.addWidget(dept_tabs_frame)
        dept_row.addStretch()
        lay.addLayout(dept_row)

        self.subdept_buttons_widget = QFrame()
        self.subdept_buttons_widget.setObjectName("subdeptPillsBar")
        self.subdept_buttons_widget.setVisible(False)
        self.subdept_btns_lay = QHBoxLayout(self.subdept_buttons_widget)
        self.subdept_btns_lay.setContentsMargins(0, 4, 0, 4)
        self.subdept_btns_lay.setSpacing(6)
        lay.addWidget(self.subdept_buttons_widget)

        # Настройки округления хвостика — под выбором отдела, видны при выборе
        self.dept_panel_host = QFrame()
        self.dept_panel_host.setFrameShape(QFrame.Shape.StyledPanel)
        self.dept_panel_host.setVisible(False)
        self.dept_panel_lay = QVBoxLayout(self.dept_panel_host)
        self.dept_panel_lay.setContentsMargins(12, 10, 12, 10)
        self.dept_panel_placeholder = QLabel("Выберите отдел для настройки округления.")
        self.dept_panel_placeholder.setObjectName("hintLabel")
        self.dept_panel_placeholder.setWordWrap(True)
        self.dept_panel_lay.addWidget(self.dept_panel_placeholder)
        lay.addWidget(self.dept_panel_host)

        main_row = QHBoxLayout()
        main_row.setSpacing(16)

        left_card = QFrame()
        left_card.setObjectName("card")
        left_card.setMinimumWidth(400)
        left_lay = QVBoxLayout(left_card)
        left_lay.setContentsMargins(16, 16, 16, 16)
        left_lay.setSpacing(12)

        products_header = QHBoxLayout()
        lbl_products = QLabel("Продукты")
        lbl_products.setObjectName("sectionTitle")
        products_header.addWidget(lbl_products)
        self.lbl_products_count = QLabel("")
        self.lbl_products_count.setObjectName("hintLabel")
        products_header.addWidget(self.lbl_products_count)
        products_header.addStretch()
        left_lay.addLayout(products_header)

        search_row = QHBoxLayout()
        search_row.addWidget(QLabel("Поиск:"))
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Введите часть названия продукта...")
        self.search_edit.textChanged.connect(self._apply_filter)
        search_row.addWidget(self.search_edit)
        left_lay.addLayout(search_row)

        self.cards_container = QWidget()
        cards_inner = QVBoxLayout(self.cards_container)
        self.cards_grid = QGridLayout()
        self.cards_grid.setSpacing(12)
        cards_inner.addLayout(self.cards_grid)
        self.lbl_no_products = QLabel("В выбранном отделе нет товаров для настройки Шт.")
        self.lbl_no_products.setObjectName("hintLabel")
        self.lbl_no_products.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_no_products.setVisible(False)
        cards_inner.addWidget(self.lbl_no_products, 1)
        left_lay.addWidget(self.cards_container, 1)

        main_row.addWidget(left_card, 1)

        lay.addLayout(main_row, 1)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_cancel = QPushButton("Отмена")
        btn_cancel.setObjectName("btnSecondary")
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_cancel)
        btn_save = QPushButton("Сохранить")
        btn_save.setObjectName("btnPrimary")
        btn_save.setDefault(True)
        btn_save.setAutoDefault(True)
        btn_save.clicked.connect(self.accept)
        btn_row.addWidget(btn_save)
        lay.addLayout(btn_row)
        QShortcut(QKeySequence(Qt.Key.Key_Return), self, self.accept)

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
                widget.hide()
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

    def _fill_data(self) -> None:
        """Заполняет кнопки отделов, подотделов и карточки продуктов."""
        self._clear_dept_buttons()
        self._clear_subdept_buttons()
        self._clear_cards()
        self._tab_defs.clear()
        self._current_tab_key = None
        self._current_scope_id = None
        self._current_scope_keys = []
        self._clear_panel_widgets(self.dept_panel_lay, keep_first=1)
        self._scope_panels.clear()
        self._product_cards.clear()

        eligible = self._eligible_products()
        self._tab_defs = self._build_tab_defs(eligible)
        if not self._tab_defs:
            self._show_dept_placeholder()
            return
        self._populate_dept_buttons()
        self._refresh_all_pcs_totals()
        self._on_dept_clicked(0)

    def _clear_dept_buttons(self) -> None:
        lay = self.dept_buttons_widget.layout()
        while lay and lay.count():
            item = lay.takeAt(0)
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
        self.subdept_buttons_widget.setVisible(False)

    def _clear_cards(self) -> None:
        while self.cards_grid.count():
            item = self.cards_grid.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

    def _populate_dept_buttons(self) -> None:
        lay = self.dept_buttons_widget.layout()
        for i, tab_def in enumerate(self._tab_defs):
            btn = QPushButton(tab_def["tab_text"])
            btn.setObjectName("deptTab")
            btn.setCheckable(True)
            btn.setChecked(i == 0)
            btn.setToolTip(f"Отдел: {tab_def['scope_label']}. Товаров: {tab_def['count']}")
            btn.clicked.connect(lambda checked, idx=i: self._on_dept_clicked(idx))
            lay.addWidget(btn)
            self._dept_buttons.append(btn)
        lay.addStretch()

    def _populate_subdept_buttons(self, tab_def: dict) -> None:
        self._clear_subdept_buttons()
        self._subdept_scopes.clear()
        sub_tabs = tab_def.get("sub_tabs", [])
        if not sub_tabs:
            self.subdept_buttons_widget.setVisible(False)
            return
        self.subdept_buttons_widget.setVisible(True)
        default_scope = tab_def["default_scope"]
        all_scopes = [default_scope] + sub_tabs
        self._subdept_scopes = all_scopes
        for i, scope_def in enumerate(all_scopes):
            btn = QPushButton(scope_def["tab_text"])
            btn.setObjectName("subdeptPill")
            btn.setCheckable(True)
            btn.setChecked(i == 0)
            btn.setToolTip(scope_def.get("scope_label", scope_def["tab_text"]))
            btn.clicked.connect(lambda c=False, s=scope_def: self._apply_scope(s))
            self.subdept_btns_lay.addWidget(btn)
            self._subdept_buttons.append(btn)
        self.subdept_btns_lay.addStretch()

    def _show_dept_placeholder(self) -> None:
        for panel in self._scope_panels.values():
            panel.setVisible(False)
        self.dept_panel_placeholder.setText("Выберите отдел для настройки округления.")
        self.dept_panel_placeholder.setVisible(True)
        self.dept_panel_host.setVisible(False)

    def _is_scope_polufabricates(self, scope_def: dict) -> bool:
        """Проверка: выбранный scope — подотдел Полуфабрикаты (по ключу или названию)."""
        dept_keys = scope_def.get("dept_keys", [])
        if any(
            excel_generator.get_dept_special_mode(k) == "polufabricates"
            for k in dept_keys
        ):
            return True
        scope_label = (scope_def.get("scope_label") or "").lower()
        return "полуфаб" in scope_label

    def _show_scope_panel(self, scope_def: dict) -> None:
        if self._is_scope_polufabricates(scope_def):
            self._clear_panel_widgets(self.dept_panel_lay, keep_first=1)
            self._scope_panels.clear()
            self.dept_panel_placeholder.setText(
                "Для подотдела «Полуфабрикаты» настройки округления (ШК, СД) не применяются — используется другая логика расчёта."
            )
            self.dept_panel_placeholder.setVisible(True)
            self.dept_panel_host.setVisible(True)
            return
        scope_id = scope_def["scope_id"]
        dept_keys = scope_def.get("dept_keys", [])
        if scope_id not in self._scope_panels:
            panel = DeptRoundPanel(
                scope_id=scope_id,
                scope_label=scope_def["scope_label"],
                dept_keys=dept_keys,
                on_changed=self._on_scope_rounding_changed,
                parent=self,
            )
            self._scope_panels[scope_id] = panel
            self.dept_panel_lay.addWidget(panel)
        self._scope_panels[scope_id]._load()
        for key, panel in self._scope_panels.items():
            panel.setVisible(key == scope_id)
        self.dept_panel_placeholder.setVisible(False)
        self.dept_panel_host.setVisible(True)

    def _apply_scope(self, scope_def: dict | None) -> None:
        self._current_scope_id = scope_def["scope_id"] if scope_def else None
        self._current_scope_keys = list(scope_def.get("dept_keys", [])) if scope_def else []
        for i, btn in enumerate(self._subdept_buttons):
            if i < len(self._subdept_scopes) and scope_def and self._subdept_scopes[i].get("scope_id") == self._current_scope_id:
                btn.setChecked(True)
            else:
                btn.setChecked(False)
        if scope_def:
            self._show_scope_panel(scope_def)
            self._populate_cards()
        else:
            self._show_dept_placeholder()
            self._clear_cards()
        self._apply_filter(self.search_edit.text())

    def _on_dept_clicked(self, index: int) -> None:
        if index < 0 or index >= len(self._tab_defs):
            return
        for i, btn in enumerate(self._dept_buttons):
            btn.setChecked(i == index)
        tab_def = self._tab_defs[index]
        self._current_tab_key = tab_def["dept_key"]
        self._populate_subdept_buttons(tab_def)
        if tab_def.get("sub_tabs"):
            self._apply_scope(tab_def["default_scope"])
        else:
            self._apply_scope(tab_def["default_scope"])

    def _populate_cards(self) -> None:
        self._clear_cards()
        self._product_cards.clear()
        products = data_store.get_ref("products") or []
        dept_by_name = {p.get("name"): p.get("deptKey") for p in products if p.get("name")}
        by_name = {p.get("name"): p for p in products if p.get("name")}
        scope_products = [
            n for n, dk in dept_by_name.items()
            if dk in self._current_scope_keys
        ]
        scope_products.sort(key=lambda x: (x or "").lower())
        cnt = len(scope_products)
        self.lbl_products_count.setText(f"({cnt} товаров)" if cnt else "(нет товаров)")
        self.lbl_no_products.setVisible(cnt == 0)
        cols = 4
        for idx, name in enumerate(scope_products):
            card = ProductCardWidget(name, self.cards_container, prod_map=by_name)
            card.clicked.connect(self._on_card_clicked)
            prod = by_name.get(name) or {}
            preview = self._card_preview_text(prod)
            card.set_preview_text(preview)
            unit = (prod.get("unit") or "").strip()
            card.set_unit_badge(unit)
            is_pcs_disabled = unit.lower() == "шт"
            card.set_disabled_for_pcs(is_pcs_disabled)
            total = self._pcs_totals_by_product.get(name, 0)
            if is_pcs_disabled:
                card.set_tooltip_text(f"{name}\nЕд. изм. «шт» — настройки округления не применяются.")
            else:
                card.set_tooltip_text(
                    f"{name}\n{preview}\n"
                    + (f"Всего по маршрутам: {_format_pcs_total(total)}" if prod.get("showPcs") else "")
                )
            self._product_cards[name] = card
            row, col = divmod(idx, cols)
            self.cards_grid.addWidget(card, row, col)

        rows = (cnt + cols - 1) // cols if cnt else 0
        card_h, gap = 120, 12
        self.cards_container.setMinimumHeight(max(200, rows * card_h + max(0, rows - 1) * gap + 40))

    def _card_preview_text(self, prod: dict) -> str:
        parts = []
        if prod.get("showPcs"):
            pcu = prod.get("pcsPerUnit") or 1
            unit = (prod.get("unit") or "").strip()
            parts.append(f"Шт: ✓ {pcu} {unit}".strip())
        else:
            parts.append("Шт: выкл.")
        if prod.get("showInDirty") and data_store.is_subdept_chistchenka(prod.get("deptKey")):
            parts.append("В Грязные ✓")
        pcs_per_unit = float(prod.get("pcsPerUnit", 1) or 1)
        shk_pct = round(_tail_value_to_percent(
            prod.get("roundTailFromШК") if prod.get("roundTailFromШК") is not None
            else (0 if prod.get("roundUpШК", True) else pcs_per_unit),
            pcs_per_unit,
        ), 0)
        sd_pct = round(_tail_value_to_percent(
            prod.get("roundTailFromСД") if prod.get("roundTailFromСД") is not None
            else (0 if prod.get("roundUpСД", True) else pcs_per_unit),
            pcs_per_unit,
        ), 0)
        parts.append(f"Хвостик: ШК {shk_pct}%, СД {sd_pct}%")
        return " | ".join(parts)

    def _on_scope_rounding_changed(self, scope_id: str) -> None:
        if scope_id in self._scope_panels:
            self._scope_panels[scope_id]._load()
        self._refresh_all_pcs_totals()
        self._refresh_card_previews()

    def _refresh_card_previews(self) -> None:
        products = data_store.get_ref("products") or []
        by_name = {p.get("name"): p for p in products if p.get("name")}
        for name, card in self._product_cards.items():
            prod = by_name.get(name) or {}
            card.set_preview_text(self._card_preview_text(prod))

    def _refresh_all_pcs_totals(self) -> None:
        self._pcs_totals_by_product = _calc_product_pcs_totals(self._app_state)
        self._refresh_card_previews()

    def _on_product_settings_changed(self, product_name: str) -> None:
        del product_name
        if self._current_scope_id and self._current_scope_id in self._scope_panels:
            self._scope_panels[self._current_scope_id]._load()
        self._refresh_all_pcs_totals()

    def _current_scope_product_list(self) -> list[str]:
        """Список продуктов текущего отдела/подотдела (отсортированный)."""
        products = data_store.get_ref("products") or []
        dept_by_name = {p.get("name"): p.get("deptKey") for p in products if p.get("name")}
        scope_products = [
            n for n, dk in dept_by_name.items()
            if dk in self._current_scope_keys
        ]
        scope_products.sort(key=lambda x: (x or "").lower())
        return scope_products

    def _on_card_clicked(self, product_name: str) -> None:
        self._open_product_pcs_dialog(product_name)

    def _open_product_pcs_dialog(self, product_name: str) -> None:
        product_list = self._current_scope_product_list()
        dlg = ProductPcsSettingsDialog(
            product_name,
            product_list,
            self._app_state,
            on_changed_callback=self._on_product_settings_changed,
            parent=self,
        )
        dlg.open_next_requested.connect(self._on_open_next_product)
        dlg.exec()

    def _on_open_next_product(self, next_product_name: str) -> None:
        sender = self.sender()
        if isinstance(sender, QDialog):
            sender.accept()
        self._open_product_pcs_dialog(next_product_name)

    def _apply_filter(self, text: str) -> None:
        pattern = (text or "").strip().lower()
        for name, card in self._product_cards.items():
            visible = not pattern or pattern in (name or "").lower()
            card.setVisible(visible)

    def _open_institutions_dialog(self) -> None:
        dlg = InstitutionsDialog(self, self._app_state)
        dlg.exec()


def open_quantity_settings_dialog(parent: QWidget, app_state: dict) -> None:
    """Открывает диалог «Настройки Количества»."""
    dlg = QuantitySettingsDialog(parent, app_state)
    dlg.exec()
