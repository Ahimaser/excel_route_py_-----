"""
settings_dialog.py — Модальное окно настроек отображения в штуках.

Открывается через open_settings_dialog(parent, app_state).
При закрытии кнопкой «Сохранить» вызывает on_saved() — коллбэк для
обновления превью-страниц.

Отличия от старого settings_page.py:
- QDialog вместо QWidget → блокирует основное окно (exec())
- Нет кнопки «Назад», есть «Сохранить» и «Отмена»
- При нажатии «Сохранить» изменения уже записаны в data_store (в реальном
  времени через чекбоксы/спинбоксы), поэтому достаточно вызвать коллбэк.
- При нажатии «Отмена» — изменения НЕ откатываются (они уже сохранены
  в data_store по мере редактирования), диалог просто закрывается.
  Это поведение аналогично старой странице.
"""
from __future__ import annotations

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTableWidget, QTableWidgetItem, QDoubleSpinBox,
    QComboBox, QHeaderView, QAbstractItemView, QLineEdit, QWidget,
    QScrollArea, QFrame,
)
from PyQt6.QtCore import Qt, QTimer

from core import data_store
from ui.widgets import hint_icon_button, ToggleSwitch


class SettingsDialog(QDialog):
    """Модальное окно настроек Шт для продуктов."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Настройки отображения в штуках")
        self.setMinimumSize(820, 560)
        self.resize(900, 620)
        # Модальность — блокирует родительское окно
        self.setWindowModality(Qt.WindowModality.ApplicationModal)

        self._updating = False
        self._search_text = ""
        self._search_timer = QTimer(self)
        self._search_timer.setSingleShot(True)
        self._search_timer.setInterval(200)
        self._search_timer.timeout.connect(self._apply_filter)

        self._build_ui()
        QTimer.singleShot(0, self._load_table)

    # ─────────────────────────── Построение UI ────────────────────────────

    def _build_ui(self):
        content = QWidget()
        lay = QVBoxLayout(content)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(14)

        # Заголовок
        title_row = QHBoxLayout()
        title_row.addWidget(QLabel("Настройки отображения в штуках"))
        title_row.addWidget(hint_icon_button(
            self,
            "Показывать Шт, кол-во в 1 шт, округление ШК/СД. Только для продуктов с ед. изм. не «шт».",
            "Инструкция — Настройки отображения в штуках\n\n"
            "1. В таблице — все продукты с единицей измерения не «шт».\n"
            "2. «Показывать Шт» — выводить количество в штуках в предпросмотре и в файлах.\n"
            "3. «Кол-во в 1 шт» — коэффициент перевода из единицы измерения в штуки.\n"
            "4. «Коэфф. замены» — множитель количества (напр. 1,25: вместо очищенных отображаются грязные овощи).\n"
            "5. «Округление ШК» и «Округление СД» — в большую или меньшую сторону.\n"
            "6. Поиск по названию — фильтрация. Изменения сохраняются автоматически.",
            "Инструкция",
        ))
        title_row.addStretch()
        lay.addLayout(title_row)

        lbl_hint = QLabel(
            "Настройки применяются ко всем продуктам с одинаковым названием. "
            "Отображение Шт — только для продуктов с ед. изм. не «шт». "
            "Коэфф. замены: отображаемое количество = количество × коэффициент (напр. 1,25 для очищенные → грязные)."
        )
        lbl_hint.setObjectName("stepLabel")
        lbl_hint.setWordWrap(True)
        lay.addWidget(lbl_hint)

        # Поиск
        search_row = QHBoxLayout()
        search_row.addWidget(QLabel("🔍 Поиск:"))
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Начните вводить название продукта...")
        self.search_edit.setClearButtonEnabled(True)
        self.search_edit.textChanged.connect(self._on_search_changed)
        search_row.addWidget(self.search_edit)
        lay.addLayout(search_row)

        # Таблица
        self.table = QTableWidget()
        self.table.setColumnCount(8)
        self.table.setHorizontalHeaderLabels([
            "Продукт", "Ед. изм.", "Показывать Шт", "Кол-во в 1 шт",
            "От кол-ва шт", "Коэфф. замены", "Хвостик ШК", "Хвостик СД"
        ])
        hdr = self.table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(6, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(7, QHeaderView.ResizeMode.ResizeToContents)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setDefaultSectionSize(80)
        lay.addWidget(self.table)

        # Кнопки
        btn_row = QHBoxLayout()
        btn_row.addStretch()

        btn_cancel = QPushButton("Закрыть")
        btn_cancel.setObjectName("btnSecondary")
        btn_cancel.setFixedHeight(36)
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_cancel)

        btn_save = QPushButton("Сохранить и закрыть")
        btn_save.setObjectName("btnPrimary")
        btn_save.setFixedHeight(36)
        btn_save.clicked.connect(self.accept)
        btn_row.addWidget(btn_save)

        lay.addLayout(btn_row)

        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setWidget(content)
        main_lay = QVBoxLayout(self)
        main_lay.setContentsMargins(0, 0, 0, 0)
        main_lay.addWidget(scroll)

    # ─────────────────────────── Поиск ────────────────────────────────────

    def _on_search_changed(self, text: str):
        self._search_text = text.strip().lower()
        self._search_timer.start()

    def _apply_filter(self):
        q = self._search_text
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            name = item.text().lower() if item else ""
            self.table.setRowHidden(row, bool(q) and q not in name)

    # ─────────────────────────── Загрузка таблицы ─────────────────────────

    def _load_table(self):
        products = data_store.get_ref("products") or []
        eligible = sorted(
            [p for p in products if p.get("unit", "").lower() != "шт"],
            key=lambda p: p["name"].lower()
        )

        self._updating = True
        self.table.setRowCount(len(eligible))

        for row, prod in enumerate(eligible):
            name = prod["name"]

            item_name = QTableWidgetItem(name)
            item_name.setFlags(item_name.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(row, 0, item_name)

            item_unit = QTableWidgetItem(prod.get("unit", ""))
            item_unit.setFlags(item_unit.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(row, 1, item_unit)

            show_pcs = prod.get("showPcs", False)

            # Переключатель «Показывать Шт»
            chk = ToggleSwitch()
            chk.setChecked(show_pcs)
            chk_widget = QWidget()
            chk_lay = QHBoxLayout(chk_widget)
            chk_lay.addWidget(chk)
            chk_lay.setAlignment(Qt.AlignmentFlag.AlignCenter)
            chk_lay.setContentsMargins(0, 0, 0, 0)
            self.table.setCellWidget(row, 2, chk_widget)

            # Спинбокс «Кол-во в 1 шт»
            spin = QDoubleSpinBox()
            spin.setRange(0.001, 99999.0)
            spin.setDecimals(3)
            spin.setSingleStep(0.1)
            spin.setValue(prod.get("pcsPerUnit", 1.0))
            spin.setEnabled(show_pcs)
            self.table.setCellWidget(row, 3, spin)

            # От какого количества считать шт (ниже — 0 шт)
            spin_min = QDoubleSpinBox()
            spin_min.setRange(0, 99999.0)
            spin_min.setDecimals(3)
            spin_min.setSingleStep(0.1)
            spin_min.setValue(prod.get("minQtyForPcs", 0) or 0)
            spin_min.setEnabled(show_pcs)
            spin_min.setToolTip("Ниже этого количества — 0 шт. От этого и выше — расчёт по «Кол-во в 1 шт» и округлению.")
            self.table.setCellWidget(row, 4, spin_min)

            # Коэфф. замены (множитель количества, напр. 1.25 для очищенные → грязные)
            spin_mult = QDoubleSpinBox()
            spin_mult.setRange(0.01, 100.0)
            spin_mult.setDecimals(2)
            spin_mult.setSingleStep(0.25)
            spin_mult.setValue(float(prod.get("quantityMultiplier", 1.0) or 1.0))
            spin_mult.setToolTip(
                "Отображаемое количество = количество × коэффициент. "
                "Пример: 1,25 для «Морковь очищенная» (кг) — в файле показывается как грязная (× 1,25)."
            )
            self.table.setCellWidget(row, 5, spin_mult)

            # Хвостик ШК: от остатка (ед.) не меньше — добавляем 1 шт
            spin_tail_shk = QDoubleSpinBox()
            spin_tail_shk.setRange(0, 99999.0)
            spin_tail_shk.setDecimals(3)
            spin_tail_shk.setSingleStep(0.01)
            v_shk = prod.get("roundTailFromШК")
            spin_tail_shk.setValue(v_shk if v_shk is not None else (0 if prod.get("roundUpШК", True) else float(prod.get("pcsPerUnit", 1))))
            spin_tail_shk.setEnabled(show_pcs)
            spin_tail_shk.setToolTip("Добавлять 1 шт, если остаток ≥ этого значения (ШК). 0 = в большую; ≥ вес 1 шт = в меньшую.")
            self.table.setCellWidget(row, 6, spin_tail_shk)

            # Хвостик СД
            spin_tail_sd = QDoubleSpinBox()
            spin_tail_sd.setRange(0, 99999.0)
            spin_tail_sd.setDecimals(3)
            spin_tail_sd.setSingleStep(0.01)
            v_sd = prod.get("roundTailFromСД")
            spin_tail_sd.setValue(v_sd if v_sd is not None else (0 if prod.get("roundUpСД", True) else float(prod.get("pcsPerUnit", 1))))
            spin_tail_sd.setEnabled(show_pcs)
            spin_tail_sd.setToolTip("Добавлять 1 шт, если остаток ≥ этого значения (СД).")
            self.table.setCellWidget(row, 7, spin_tail_sd)

            # Сигналы
            chk.stateChanged.connect(
                lambda state, n=name, s=spin, sm=spin_min, t1=spin_tail_shk, t2=spin_tail_sd: self._on_show_pcs(n, state, s, sm, t1, t2)
            )
            spin.valueChanged.connect(lambda val, n=name: self._on_pcs_per_unit(n, val))
            spin_min.valueChanged.connect(lambda val, n=name: self._on_min_qty_for_pcs(n, val))
            spin_mult.valueChanged.connect(lambda val, n=name: self._on_multiplier(n, val))
            spin_tail_shk.valueChanged.connect(lambda val, n=name: self._on_tail_shk(n, val))
            spin_tail_sd.valueChanged.connect(lambda val, n=name: self._on_tail_sd(n, val))

        self._updating = False

    # ─────────────────────────── Обработчики изменений ────────────────────

    def _on_show_pcs(self, name: str, state: int, spin: QDoubleSpinBox,
                     spin_min: QDoubleSpinBox, spin_tail_shk: QDoubleSpinBox, spin_tail_sd: QDoubleSpinBox):
        if self._updating:
            return
        show = state == Qt.CheckState.Checked.value
        spin.setEnabled(show)
        spin_min.setEnabled(show)
        spin_tail_shk.setEnabled(show)
        spin_tail_sd.setEnabled(show)
        data_store.update_product(name, showPcs=show)

    def _on_pcs_per_unit(self, name: str, val: float):
        if self._updating:
            return
        data_store.update_product(name, pcsPerUnit=val)

    def _on_min_qty_for_pcs(self, name: str, val: float):
        if self._updating:
            return
        data_store.update_product(name, minQtyForPcs=val if val > 0 else None)

    def _on_tail_shk(self, name: str, val: float):
        if self._updating:
            return
        data_store.update_product(name, roundTailFromШК=val)

    def _on_tail_sd(self, name: str, val: float):
        if self._updating:
            return
        data_store.update_product(name, roundTailFromСД=val)

    def _on_multiplier(self, name: str, val: float):
        if self._updating:
            return
        data_store.update_product(name, quantityMultiplier=val if val != 1.0 else None)


# ─────────────────────────── Диалог «Кол-во в шт» для одного продукта ──────

class PcsSettingsDialog(QDialog):
    """Модальное окно настроек отображения в штуках для одного продукта."""

    def __init__(self, product_name: str, parent=None, on_saved=None):
        super().__init__(parent)
        self._product_name = product_name
        self._on_saved = on_saved
        self.setWindowTitle("Кол-во в шт")
        self.setWindowModality(Qt.WindowModality.ApplicationModal)
        self.setMinimumSize(420, 320)
        self._build_ui()
        self._load_product()

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(16)

        self.lbl_product = QLabel("")
        self.lbl_product.setObjectName("sectionTitle")
        lay.addWidget(self.lbl_product)

        self.lbl_hint = QLabel(
            "Настройка применяется только к продуктам с единицей измерения, отличной от «шт»."
        )
        self.lbl_hint.setObjectName("stepLabel")
        self.lbl_hint.setWordWrap(True)
        lay.addWidget(self.lbl_hint)

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

        # От какого количества считать шт
        row_min = QHBoxLayout()
        row_min.addWidget(QLabel("От кол-ва шт:"))
        self.spin_min_qty = QDoubleSpinBox()
        self.spin_min_qty.setRange(0, 99999.0)
        self.spin_min_qty.setDecimals(3)
        self.spin_min_qty.setSingleStep(0.1)
        self.spin_min_qty.setToolTip("Ниже этого количества — 0 шт. От этого и выше — расчёт по «Кол-во в 1 шт» и округлению.")
        self.spin_min_qty.valueChanged.connect(self._on_min_qty_changed)
        row_min.addWidget(self.spin_min_qty)
        row_min.addStretch()
        lay.addLayout(row_min)

        # Коэфф. замены с пояснением и примером
        row_mult = QHBoxLayout()
        row_mult.addWidget(QLabel("Коэфф. замены (множитель количества):"))
        self.spin_mult = QDoubleSpinBox()
        self.spin_mult.setRange(0.01, 100.0)
        self.spin_mult.setDecimals(2)
        self.spin_mult.setSingleStep(0.25)
        _mult_tip = (
            "Отображаемое количество = количество × коэффициент.\n"
            "Пример: для «Морковь очищенная» (кг) коэффициент 1,25 — в файле показывается количество × 1,25 "
            "(как грязная морковь)."
        )
        self.spin_mult.setToolTip(_mult_tip)
        self.spin_mult.valueChanged.connect(self._on_multiplier_changed)
        row_mult.addWidget(self.spin_mult)
        row_mult.addStretch()
        lay.addLayout(row_mult)
        lbl_mult_hint = QLabel(
            "Пример: 1,25 для «очищенная» (кг) — в файле отображается количество × 1,25 (как грязная)."
        )
        lbl_mult_hint.setObjectName("hintLabel")
        lbl_mult_hint.setWordWrap(True)
        lay.addWidget(lbl_mult_hint)

        # Хвостик ШК — граница остатка: при остатке от неё включительно округление в большую
        row3 = QHBoxLayout()
        row3.addWidget(QLabel("Хвостик ШК (остаток ≥ — в большую):"))
        self.spin_tail_shk = QDoubleSpinBox()
        self.spin_tail_shk.setRange(0, 99999.0)
        self.spin_tail_shk.setDecimals(3)
        self.spin_tail_shk.setSingleStep(0.01)
        self.spin_tail_shk.setToolTip(
            "Граница остатка (кг, л и т.д.): при остатке от этого значения включительно — +1 шт. "
            "Пример: 1 шт = 0.7, хвостик = 0.35 → при остатке ≥ 0.35 в большую. 0 = при любом остатке в большую."
        )
        self.spin_tail_shk.valueChanged.connect(lambda: self._on_tail_shk(self.spin_tail_shk.value()))
        row3.addWidget(self.spin_tail_shk)
        row3.addStretch()
        lay.addLayout(row3)

        # Хвостик СД — граница остатка: при остатке от неё включительно округление в большую
        row4 = QHBoxLayout()
        row4.addWidget(QLabel("Хвостик СД (остаток ≥ — в большую):"))
        self.spin_tail_sd = QDoubleSpinBox()
        self.spin_tail_sd.setRange(0, 99999.0)
        self.spin_tail_sd.setDecimals(3)
        self.spin_tail_sd.setSingleStep(0.01)
        self.spin_tail_sd.setToolTip(
            "Граница остатка (кг, л и т.д.): при остатке от этого значения включительно — +1 шт. 0 = при любом остатке в большую."
        )
        self.spin_tail_sd.valueChanged.connect(lambda: self._on_tail_sd(self.spin_tail_sd.value()))
        row4.addWidget(self.spin_tail_sd)
        row4.addStretch()
        lay.addLayout(row4)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_cancel = QPushButton("Закрыть")
        btn_cancel.setObjectName("btnSecondary")
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_cancel)
        btn_ok = QPushButton("Сохранить")
        btn_ok.setObjectName("btnPrimary")
        btn_ok.clicked.connect(self._save_and_close)
        btn_row.addWidget(btn_ok)
        lay.addLayout(btn_row)

    def _load_product(self):
        products = data_store.get_ref("products") or []
        prod = next((p for p in products if p.get("name") == self._product_name), None)
        if not prod:
            self.lbl_product.setText(self._product_name)
            self.chk_show.setEnabled(False)
            self.spin_pcs.setEnabled(False)
            self.spin_min_qty.setEnabled(False)
            self.spin_tail_shk.setEnabled(False)
            self.spin_tail_sd.setEnabled(False)
            return
        unit = (prod.get("unit") or "").strip().lower()
        self.lbl_product.setText(f"{prod['name']}  ({prod.get('unit', '')})")
        if unit == "шт":
            self.lbl_hint.setText(
                "Для продуктов с единицей измерения «шт» настройка «Кол-во в шт» не применяется."
            )
            self.chk_show.setEnabled(False)
            self.spin_pcs.setEnabled(False)
            self.spin_min_qty.setEnabled(False)
            self.spin_tail_shk.setEnabled(False)
            self.spin_tail_sd.setEnabled(False)
            return
        show_pcs = prod.get("showPcs", False)
        self.chk_show.setChecked(show_pcs)
        self.spin_pcs.setValue(prod.get("pcsPerUnit", 1.0))
        self.spin_min_qty.setValue(prod.get("minQtyForPcs", 0) or 0)
        v_shk = prod.get("roundTailFromШК")
        self.spin_tail_shk.setValue(v_shk if v_shk is not None else (0 if prod.get("roundUpШК", True) else float(prod.get("pcsPerUnit", 1))))
        v_sd = prod.get("roundTailFromСД")
        self.spin_tail_sd.setValue(v_sd if v_sd is not None else (0 if prod.get("roundUpСД", True) else float(prod.get("pcsPerUnit", 1))))
        self.spin_pcs.setEnabled(show_pcs)
        self.spin_min_qty.setEnabled(show_pcs)
        self.spin_tail_shk.setEnabled(show_pcs)
        self.spin_tail_sd.setEnabled(show_pcs)
        self.spin_mult.setValue(float(prod.get("quantityMultiplier", 1.0) or 1.0))

    def _on_show_changed(self, state: int):
        show = state == Qt.CheckState.Checked.value
        self.spin_pcs.setEnabled(show)
        self.spin_min_qty.setEnabled(show)
        self.spin_tail_shk.setEnabled(show)
        self.spin_tail_sd.setEnabled(show)
        data_store.update_product(self._product_name, showPcs=show)

    def _on_pcs_changed(self, val: float):
        data_store.update_product(self._product_name, pcsPerUnit=val)

    def _on_min_qty_changed(self, val: float):
        data_store.update_product(self._product_name, minQtyForPcs=val if val > 0 else None)

    def _on_tail_shk(self, val: float):
        data_store.update_product(self._product_name, roundTailFromШК=val)

    def _on_tail_sd(self, val: float):
        data_store.update_product(self._product_name, roundTailFromСД=val)

    def _on_multiplier_changed(self, val: float):
        data_store.update_product(
            self._product_name,
            quantityMultiplier=val if val != 1.0 else None,
        )

    def _save_and_close(self):
        if self._on_saved is not None:
            self._on_saved()
        self.accept()


def open_pcs_for_product(parent, product_name: str, on_saved=None):
    """
    Открывает диалог настроек «Кол-во в шт» для одного продукта.

    Args:
        parent:     родительский виджет
        product_name: название продукта
        on_saved:   коллбэк при сохранении (опционально)
    """
    dlg = PcsSettingsDialog(product_name, parent, on_saved=on_saved)
    dlg.exec()


# ─────────────────────────── Публичная функция (старый диалог — все продукты) ─

def open_settings_dialog(parent, on_saved=None):
    """
    Открывает модальное окно настроек Шт (все продукты).
    Используется редко; основной доступ — Справочник продуктов → ПКМ по продукту → «Кол-во в шт».
    """
    dlg = SettingsDialog(parent)
    result = dlg.exec()
    if result == QDialog.DialogCode.Accepted and on_saved is not None:
        on_saved()
