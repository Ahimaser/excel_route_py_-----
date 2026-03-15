"""
product_replacement_dialog.py — Диалог замены продукта.

Кнопка «Замена продукта» на странице общие маршруты.
- Выбор продукта для замены → общее количество по маршрутам
- Режим: целиком / частично (X + выбор учреждений)
- При частичной замене: списание с последних по порядку маршрутов
- Замены отображаются во всех сохраняемых файлах
"""
from __future__ import annotations

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QShortcut, QKeySequence
from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QComboBox,
    QRadioButton,
    QDoubleSpinBox,
    QScrollArea,
    QFrame,
    QCheckBox,
    QButtonGroup,
    QGroupBox,
    QMessageBox,
)

from core import data_store
from ui.widgets import ReplacementDiagramWidget


def _get_product_totals(routes: list) -> dict[str, tuple[float, str]]:
    """Возвращает {product_name: (total_quantity, unit)} по маршрутам."""
    result: dict[str, tuple[float, str]] = {}
    for r in routes or []:
        if r.get("excluded"):
            continue
        for p in r.get("products", []):
            name = p.get("name", "")
            if not name:
                continue
            try:
                qty = float(p.get("quantity") or 0)
            except (TypeError, ValueError):
                qty = 0
            unit = (p.get("unit") or "").strip()
            prev_qty, prev_unit = result.get(name, (0, unit))
            result[name] = (prev_qty + qty, prev_unit or unit)
    return result


def _get_institutions_with_product(routes: list, product_name: str) -> dict[str, list[str]]:
    """Учреждения и адреса, в маршрутах которых есть product_name. {code: [addresses]}."""
    inst_map = data_store.get_institution_addresses_map(routes or [])
    result: dict[str, list[str]] = {}
    for r in routes or []:
        if r.get("excluded"):
            continue
        if not any(p.get("name") == product_name for p in r.get("products", [])):
            continue
        addr = (r.get("address") or "").strip()
        if not addr:
            continue
        key = data_store.get_institution_key_from_address(addr)
        if key and key in inst_map:
            result[key] = inst_map[key]
    return result


def _get_product_qty_by_address(routes: list, product_name: str) -> dict[str, float]:
    """Количество продукта по адресам. {address: quantity}."""
    result: dict[str, float] = {}
    for r in routes or []:
        if r.get("excluded"):
            continue
        addr = (r.get("address") or "").strip()
        if not addr:
            continue
        for p in r.get("products", []):
            if p.get("name") == product_name:
                try:
                    qty = float(p.get("quantity") or 0)
                except (TypeError, ValueError):
                    qty = 0
                result[addr] = result.get(addr, 0) + qty
                break
    return result


def open_product_replacement_dialog(parent, app_state: dict) -> None:
    """Открывает диалог замены продукта. Результат пишется в app_state['productReplacements']."""
    dlg = ProductReplacementDialog(parent, app_state)
    dlg.exec()


class ProductReplacementDialog(QDialog):
    """Диалог настройки замены продукта."""

    def __init__(self, parent, app_state: dict):
        super().__init__(parent)
        self.app_state = app_state
        self.setWindowTitle("Замена продукта")
        self.setMinimumWidth(640)
        self.setMinimumHeight(620)
        self.resize(680, 680)

        self._routes = [
            r for r in (app_state.get("filteredRoutes") or app_state.get("routes") or [])
            if not r.get("excluded")
        ]
        self._replacements: list[dict] = list(app_state.get("productReplacements") or [])

        self._build_ui()
        self._refresh_product_info()

    def _build_ui(self) -> None:
        lay = QVBoxLayout(self)
        lay.setSpacing(12)

        # Визуальная схема замены (вариант 2)
        lay.addWidget(QLabel("Схема замены:"))
        self.diagram = ReplacementDiagramWidget(self)
        self.diagram.combo_from.currentIndexChanged.connect(self._on_from_changed)
        self.diagram.combo_to1.currentIndexChanged.connect(self._on_to1_changed)
        self.diagram.btn_add_second.toggled.connect(self._on_second_product_toggled)
        self.diagram.slider_ratio.valueChanged.connect(lambda: self._on_ratio_changed())
        lay.addWidget(self.diagram)

        self.lbl_remaining = QLabel("")
        self.lbl_remaining.setObjectName("hintLabel")
        lay.addWidget(self.lbl_remaining)

        self.lbl_empty_hint = QLabel("")
        self.lbl_empty_hint.setObjectName("emptyHintLabel")
        self.lbl_empty_hint.setWordWrap(True)
        lay.addWidget(self.lbl_empty_hint)

        # Режим
        grp_mode = QGroupBox("Режим замены")
        mode_lay = QVBoxLayout(grp_mode)
        self.btn_full = QRadioButton("Заменить целиком во всех маршрутах")
        self.btn_partial = QRadioButton("Заменить частично:")
        self.btn_full.setChecked(True)
        self.btn_group = QButtonGroup(self)
        self.btn_group.addButton(self.btn_full)
        self.btn_group.addButton(self.btn_partial)
        mode_lay.addWidget(self.btn_full)
        partial_row = QHBoxLayout()
        partial_row.addWidget(self.btn_partial)
        self.spin_qty = QDoubleSpinBox()
        self.spin_qty.setRange(0.001, 999999)
        self.spin_qty.setDecimals(2)
        self.spin_qty.setMinimumWidth(100)
        self.spin_qty.valueChanged.connect(self._on_partial_qty_changed)
        partial_row.addWidget(self.spin_qty)
        self.lbl_unit_spin = QLabel("")
        partial_row.addWidget(self.lbl_unit_spin)
        partial_row.addStretch()
        mode_lay.addLayout(partial_row)
        lay.addWidget(grp_mode)

        # Учреждения (при частичной замене)
        self.grp_inst = QGroupBox("Учреждения для замены")
        inst_lay = QVBoxLayout(self.grp_inst)
        self.lbl_inst_hint = QLabel(
            "Выберите учреждения целиком или отдельные адреса. "
            "При частичной замене списание производится с последних по порядку маршрутов."
        )
        self.lbl_inst_hint.setObjectName("stepLabel")
        self.lbl_inst_hint.setWordWrap(True)
        inst_lay.addWidget(self.lbl_inst_hint)
        btn_row = QHBoxLayout()
        self.btn_select_all = QPushButton("Выбрать все")
        self.btn_select_all.setObjectName("btnSecondary")
        self.btn_select_all.clicked.connect(lambda: self._set_all_institutions(True))
        self.btn_deselect_all = QPushButton("Снять все")
        self.btn_deselect_all.setObjectName("btnSecondary")
        self.btn_deselect_all.clicked.connect(lambda: self._set_all_institutions(False))
        btn_row.addWidget(self.btn_select_all)
        btn_row.addWidget(self.btn_deselect_all)
        btn_row.addStretch()
        inst_lay.addLayout(btn_row)
        self.lbl_selected_total = QLabel("")
        self.lbl_selected_total.setObjectName("selectedTotalLabel")
        inst_lay.addWidget(self.lbl_selected_total)
        self.inst_scroll = QScrollArea()
        self.inst_scroll.setWidgetResizable(True)
        self.inst_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.inst_scroll.setMinimumHeight(220)
        self.inst_container = QFrame()
        self.inst_lay = QVBoxLayout(self.inst_container)
        self.inst_lay.setContentsMargins(0, 0, 0, 0)
        self.inst_scroll.setWidget(self.inst_container)
        inst_lay.addWidget(self.inst_scroll, 1)
        lay.addWidget(self.grp_inst)
        self.grp_inst.setVisible(False)
        def _on_mode_toggled(on: bool) -> None:
            self.grp_inst.setVisible(on)
            self._update_remaining_label()
        self.btn_partial.toggled.connect(_on_mode_toggled)

        # Активные замены
        self.lbl_active = QLabel("")
        self.lbl_active.setObjectName("hintLabel")
        self.lbl_active.setWordWrap(True)
        lay.addWidget(self.lbl_active)

        # Кнопки
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.btn_clear = QPushButton("Очистить замены")
        self.btn_clear.setObjectName("btnSecondary")
        self.btn_clear.setToolTip("Ctrl+Delete")
        self.btn_clear.clicked.connect(self._on_clear)
        btn_row.addWidget(self.btn_clear)
        self.btn_apply = QPushButton("Применить")
        self.btn_apply.setObjectName("btnPrimary")
        self.btn_apply.setToolTip("Ctrl+Enter")
        self.btn_apply.clicked.connect(self._on_apply)
        btn_row.addWidget(self.btn_apply)
        btn_cancel = QPushButton("Отмена")
        btn_cancel.setObjectName("btnSecondary")
        btn_cancel.setToolTip("Esc")
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_cancel)
        lay.addLayout(btn_row)

        # Горячие клавиши
        QShortcut(QKeySequence("Ctrl+Return"), self, self._on_apply)
        QShortcut(QKeySequence("Ctrl+Enter"), self, self._on_apply)
        QShortcut(QKeySequence(Qt.Key.Key_Escape), self, self.reject)
        QShortcut(QKeySequence("Ctrl+Delete"), self, self._on_clear)

        self._populate_products()

    def _populate_products(self) -> None:
        totals = _get_product_totals(self._routes)
        from_names = sorted(totals.keys())
        self.diagram.combo_from.blockSignals(True)
        self.diagram.combo_from.clear()
        self.diagram.combo_from.addItem("— Выберите продукт —", "")
        for n in from_names:
            self.diagram.combo_from.addItem(n, n)
        self.diagram.combo_from.blockSignals(False)
        self._refresh_product_info()

    def _refresh_product_info(self) -> None:
        from_name = self.diagram.combo_from.currentData() or ""
        totals = _get_product_totals(self._routes)

        # Подсказка при пустом списке продуктов
        if not totals:
            self.lbl_empty_hint.setText(
                "Нет продуктов в маршрутах. Загрузите файлы маршрутов на странице «Обработка файлов»."
            )
        else:
            self.lbl_empty_hint.setText("")

        if from_name:
            qty, unit = totals.get(from_name, (0, ""))
            self.diagram.lbl_from_qty.setText(f"{qty} {unit}".strip())
            self.lbl_unit_spin.setText(unit)
            self.spin_qty.setMaximum(max(qty, 999999))
            if qty > 0:
                self.spin_qty.setValue(min(self.spin_qty.value(), qty))
            self._populate_replace_with(from_name)
            self._populate_institutions(from_name, totals)
            self._load_replacement_for_edit(from_name, totals)
        else:
            self.diagram.lbl_from_qty.setText("")
            self.lbl_remaining.setText("")
            self.lbl_unit_spin.setText("")
            self._populate_replace_with("")
            self._clear_institutions()
        self._update_remaining_label(totals)
        self._update_active_label()

    def _populate_replace_with(self, from_name: str) -> None:
        """Заполняет списки «Заменить на»: только продукты того же отдела, исключая from_name."""
        products = data_store.get_ref("products") or []
        from_dept = None
        if from_name:
            from_dept = next(
                (p.get("deptKey") for p in products if p.get("name") == from_name),
                None
            )
        to_names = [
            p.get("name") for p in products
            if p.get("name")
            and p.get("name") != from_name
            and (from_dept is None and p.get("deptKey") is None
                 or from_dept is not None and p.get("deptKey") == from_dept)
        ]
        to_names = sorted(set(to_names))
        self.diagram.combo_to1.blockSignals(True)
        self.diagram.combo_to1.clear()
        self.diagram.combo_to1.addItem("— Выберите продукт —", "")
        for n in to_names:
            self.diagram.combo_to1.addItem(n, n)
        self.diagram.combo_to1.blockSignals(False)
        to1 = self.diagram.combo_to1.currentData() or ""
        self.diagram.combo_to2.blockSignals(True)
        self.diagram.combo_to2.clear()
        self.diagram.combo_to2.addItem("— Выберите продукт —", "")
        for n in to_names:
            if n != to1:
                self.diagram.combo_to2.addItem(n, n)
        self.diagram.combo_to2.blockSignals(False)

    def _on_from_changed(self) -> None:
        self._refresh_product_info()

    def _on_second_product_toggled(self, checked: bool) -> None:
        if checked:
            self._populate_replace_with(self.diagram.combo_from.currentData() or "")
        self._on_ratio_changed()

    def _on_ratio_changed(self) -> None:
        if self.diagram.btn_add_second.isChecked():
            r = self.diagram.slider_ratio.value()
            self.diagram.lbl_to1_pct.setText(f"{r}%")
            self.diagram.lbl_to2_pct.setText(f"{100 - r}%")

    def _load_replacement_for_edit(self, from_name: str, totals: dict) -> None:
        """Подставляет параметры активной замены в форму при выборе продукта."""
        repl = next((r for r in self._replacements if r.get("fromProduct") == from_name), None)
        if not repl:
            return
        self.diagram.combo_to1.blockSignals(True)
        idx = self.diagram.combo_to1.findData(repl.get("toProduct") or "")
        self.diagram.combo_to1.setCurrentIndex(max(0, idx))
        self.diagram.combo_to1.blockSignals(False)

        to_products = repl.get("toProducts")
        if to_products and len(to_products) >= 2:
            self.diagram.btn_add_second.setChecked(True)
            self.diagram.combo_to2.blockSignals(True)
            idx2 = self.diagram.combo_to2.findData(to_products[1])
            self.diagram.combo_to2.setCurrentIndex(max(0, idx2))
            self.diagram.combo_to2.blockSignals(False)
            r = int((repl.get("splitRatio") or 0.5) * 100)
            self.diagram.slider_ratio.setValue(r)
        else:
            self.diagram.btn_add_second.setChecked(False)

        if repl.get("mode") == "partial":
            self.btn_partial.setChecked(True)
            self.spin_qty.setValue(float(repl.get("quantity") or 0))
            addrs = set(repl.get("addresses") or [])
            if not addrs and repl.get("institutionCodes"):
                inst_map = _get_institutions_with_product(self._routes, from_name)
                for code in repl.get("institutionCodes", []):
                    addrs.update(inst_map.get(code, []))
            for addr, chk in getattr(self, "_addr_checks", {}).items():
                chk.blockSignals(True)
                chk.setChecked(addr in addrs)
                chk.blockSignals(False)
            for code, chk in getattr(self, "_inst_checks", {}).items():
                addrs_for_code = [a for a, c in self._addr_checks.items() if c.property("code") == code]
                chk.blockSignals(True)
                chk.setChecked(all(a in addrs for a in addrs_for_code))
                chk.blockSignals(False)
        else:
            self.btn_full.setChecked(True)
        self._update_remaining_label(totals)
        self._update_selected_total(totals)

    def _on_to1_changed(self) -> None:
        if self.diagram.btn_add_second.isChecked():
            self._populate_replace_with(self.diagram.combo_from.currentData() or "")

    def _on_partial_qty_changed(self) -> None:
        self._update_remaining_label()
        self._update_active_label()

    def _update_remaining_label(self, totals: dict | None = None) -> None:
        """Обновляет строку «Осталось заменить X ед. изм. продукта»."""
        from_name = self.diagram.combo_from.currentData() or ""
        if totals is None:
            totals = _get_product_totals(self._routes)
        if not from_name:
            self.lbl_remaining.setText("")
            return
        qty, unit = totals.get(from_name, (0, ""))
        unit = (unit or "").strip()
        if self.btn_full.isChecked():
            self.lbl_remaining.setText(f"Осталось заменить: 0 {unit}".strip())
            return
        to_replace = self.spin_qty.value()
        remaining = max(0, qty - to_replace)
        self.lbl_remaining.setText(f"Осталось заменить: {remaining} {unit}".strip())

    def _populate_institutions(self, product_name: str, totals: dict | None = None) -> None:
        for i in reversed(range(self.inst_lay.count())):
            w = self.inst_lay.takeAt(i).widget()
            if w:
                w.deleteLater()
        inst_map = _get_institutions_with_product(self._routes, product_name)
        self._inst_checks: dict[str, QCheckBox] = {}
        self._addr_checks: dict[str, QCheckBox] = {}
        for code in sorted(inst_map.keys()):
            addrs = inst_map[code]
            chk_inst = QCheckBox(f"{code} — всё учреждение ({len(addrs)} адр.)")
            chk_inst.setProperty("code", code)
            chk_inst.setObjectName("boldCheckBox")
            def _make_inst_handler(c):
                def _h():
                    self._on_inst_toggled(c, self._inst_checks[c].isChecked())
                return _h
            chk_inst.stateChanged.connect(_make_inst_handler(code))
            chk_inst.stateChanged.connect(self._update_selected_total)
            self._inst_checks[code] = chk_inst
            self.inst_lay.addWidget(chk_inst)
            for addr in addrs:
                chk_addr = QCheckBox(addr)
                chk_addr.setProperty("address", addr)
                chk_addr.setProperty("code", code)
                chk_addr.setObjectName("indentedCheckBox")
                chk_addr.stateChanged.connect(lambda c=code: self._sync_inst_check(c))
                chk_addr.stateChanged.connect(self._update_selected_total)
                self._addr_checks[addr] = chk_addr
                self.inst_lay.addWidget(chk_addr)
        self._update_selected_total(totals)

    def _on_inst_toggled(self, code: str, checked: bool) -> None:
        for addr, chk in getattr(self, "_addr_checks", {}).items():
            if chk.property("code") == code:
                chk.blockSignals(True)
                chk.setChecked(checked)
                chk.blockSignals(False)
        self._update_selected_total()

    def _update_selected_total(self, totals: dict | None = None) -> None:
        """Обновляет метку «В выбранных: X ед.»."""
        if not hasattr(self, "lbl_selected_total"):
            return
        from_name = self.diagram.combo_from.currentData() or ""
        if not from_name:
            self.lbl_selected_total.setText("")
            return
        qty_by_addr = _get_product_qty_by_address(self._routes, from_name)
        selected = self._get_selected_addresses()
        total = sum(qty_by_addr.get(addr, 0) for addr in selected)
        if totals is None:
            totals = _get_product_totals(self._routes)
        _, unit = totals.get(from_name, (0, ""))
        unit = (unit or "").strip()
        if selected:
            try:
                if abs(total - round(total)) < 1e-9:
                    fmt = str(int(round(total)))
                else:
                    fmt = f"{total:.2f}".rstrip("0").rstrip(".")
            except (TypeError, ValueError):
                fmt = str(total)
            self.lbl_selected_total.setText(f"В выбранных учреждениях: {fmt} {unit}".strip())
        else:
            self.lbl_selected_total.setText("Выберите учреждения или адреса")

    def _sync_inst_check(self, code: str) -> None:
        addrs_for_code = [a for a, chk in getattr(self, "_addr_checks", {}).items() if chk.property("code") == code]
        checked = sum(1 for a in addrs_for_code if self._addr_checks[a].isChecked())
        inst_chk = self._inst_checks.get(code)
        if inst_chk:
            inst_chk.blockSignals(True)
            inst_chk.setChecked(checked == len(addrs_for_code))
            inst_chk.blockSignals(False)

    def _clear_institutions(self) -> None:
        for i in reversed(range(self.inst_lay.count())):
            w = self.inst_lay.takeAt(i).widget()
            if w:
                w.deleteLater()
        self._inst_checks = {}
        self._addr_checks = {}
        if hasattr(self, "lbl_selected_total"):
            self.lbl_selected_total.setText("")

    def _set_all_institutions(self, checked: bool) -> None:
        for chk in getattr(self, "_addr_checks", {}).values():
            chk.setChecked(checked)
        for chk in getattr(self, "_inst_checks", {}).values():
            chk.setChecked(checked)
        self._update_selected_total()

    def _check_show_pcs_sync(self, from_name: str, to_name: str) -> bool:
        """Проверяет синхронизацию showPcs: оба продукта должны иметь одинаковую настройку."""
        products = data_store.get_ref("products") or []
        pm = {p.get("name"): p for p in products if p.get("name")}
        sp_from = pm.get(from_name, {})
        sp_to = pm.get(to_name, {})
        show_from = bool(sp_from.get("showPcs"))
        show_to = bool(sp_to.get("showPcs"))
        if show_from != show_to:
            msg = (
                f"Для корректного отображения Шт оба продукта должны иметь одинаковую настройку «Показывать Шт».\n\n"
                f"«{from_name}»: {'включено' if show_from else 'выключено'}\n"
                f"«{to_name}»: {'включено' if show_to else 'выключено'}\n\n"
                "Откройте «Настройки» → «Настройки Количества» и включите/выключите Шт для обоих продуктов."
            )
            QMessageBox.warning(self, "Настройки Шт", msg)
            return False
        return True

    def _get_selected_addresses(self) -> list[str]:
        return [
            addr for addr, chk in getattr(self, "_addr_checks", {}).items()
            if chk.isChecked()
        ]

    def _update_active_label(self) -> None:
        if not self._replacements:
            self.lbl_active.setText("")
            return
        parts = []
        for r in self._replacements:
            to_ps = r.get("toProducts")
            if to_ps:
                s = f"{r['fromProduct']} → {to_ps[0]} + {to_ps[1]}"
                if r.get("splitRatio") is not None:
                    pct = int((r.get("splitRatio") or 0.5) * 100)
                    s += f" ({pct}% / {100 - pct}%)"
            else:
                s = f"{r['fromProduct']} → {r['toProduct']}"
            if r.get("mode") == "partial":
                addrs = r.get("addresses") or r.get("institutionCodes") or []
                s += f" ({r.get('quantity', 0)} {r.get('unit', '')} в {len(addrs)} адр.)"
            parts.append(s)
        self.lbl_active.setText("Активные замены:\n• " + "\n• ".join(parts))

    def _on_clear(self) -> None:
        self._replacements = []
        self.app_state["productReplacements"] = []
        self._update_active_label()
        self._populate_products()

    def _on_apply(self) -> None:
        from_name = self.diagram.combo_from.currentData() or ""
        to_name = self.diagram.combo_to1.currentData() or ""
        if not from_name or not to_name:
            QMessageBox.information(
                self,
                "Замена продукта",
                "Выберите продукт для замены и продукт(ы), на которые нужно заменить.",
            )
            return
        if from_name == to_name:
            return
        use_two = self.diagram.btn_add_second.isChecked()
        to_name2 = self.diagram.combo_to2.currentData() or "" if use_two else ""
        if use_two and (not to_name2 or to_name2 == to_name or to_name2 == from_name):
            QMessageBox.warning(self, "Второй продукт", "Выберите второй продукт, отличный от первого.")
            return
        if not self._check_show_pcs_sync(from_name, to_name):
            return
        if use_two and not self._check_show_pcs_sync(from_name, to_name2):
            return
        totals = _get_product_totals(self._routes)
        qty, unit = totals.get(from_name, (0, ""))

        mode = "full" if self.btn_full.isChecked() else "partial"
        repl: dict = {
            "fromProduct": from_name,
            "toProduct": to_name,
            "mode": mode,
            "unit": unit,
        }
        if use_two:
            repl["toProducts"] = [to_name, to_name2]
            repl["splitRatio"] = self.diagram.slider_ratio.value() / 100.0  # доля на первый (0..1)
        if mode == "partial":
            repl["quantity"] = self.spin_qty.value()
            addrs = self._get_selected_addresses()
            if not addrs:
                QMessageBox.warning(
                    self, "Выбор адресов",
                    "При частичной замене выберите хотя бы одно учреждение или адрес."
                )
                return
            repl["addresses"] = addrs

        # Удаляем старую замену для того же fromProduct
        self._replacements = [r for r in self._replacements if r.get("fromProduct") != from_name]
        self._replacements.append(repl)
        self.app_state["productReplacements"] = list(self._replacements)
        self._update_active_label()
