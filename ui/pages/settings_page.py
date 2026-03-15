"""
settings_page.py — Настройки отображения в штуках для каждого продукта.

Оптимизации:
- Кэш продуктов в памяти (_products_cache) — не читаем JSON при каждом изменении
- Используем data_store.update_product() для точечного обновления одного продукта
- Таблица перестраивается только при refresh() (не при каждом изменении виджета)
- Виджеты ячеек хранят ссылку на имя продукта через замыкание (без поиска по тексту)
"""
from __future__ import annotations

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFrame, QTableWidget, QTableWidgetItem,
    QDoubleSpinBox, QComboBox, QHeaderView,
    QAbstractItemView, QLineEdit, QScrollArea
)
from PyQt6.QtCore import Qt, pyqtSignal, QTimer

from core import data_store
from ui.widgets import ToggleSwitch


class SettingsPage(QWidget):
    """Страница настроек Шт для продуктов."""

    go_back = pyqtSignal()

    def __init__(self, app_state: dict = None):
        super().__init__()
        self.app_state = app_state or {}
        self._updating = False
        self._search_text = ""
        self._search_timer = QTimer(self)
        self._search_timer.setSingleShot(True)
        self._search_timer.setInterval(200)
        self._search_timer.timeout.connect(self._apply_filter)
        self._build_ui()

    def _build_ui(self):
        content = QWidget()
        lay = QVBoxLayout(content)
        lay.setContentsMargins(32, 24, 32, 24)
        lay.setSpacing(16)

        h_row = QHBoxLayout()
        btn_back = QPushButton("← Назад")
        btn_back.setObjectName("btnBack")
        btn_back.clicked.connect(self.go_back.emit)
        h_row.addWidget(btn_back)

        lbl = QLabel("Настройки отображения в штуках")
        lbl.setObjectName("sectionTitle")
        h_row.addWidget(lbl)
        h_row.addStretch()
        lay.addLayout(h_row)

        lbl_hint = QLabel(
            "Настройки применяются ко всем продуктам с одинаковым названием. "
            "Отображение в штуках доступно только для продуктов с единицей измерения, отличной от «шт»."
        )
        lbl_hint.setObjectName("stepLabel")
        lbl_hint.setWordWrap(True)
        lay.addWidget(lbl_hint)

        # Поиск по названию продукта
        search_row = QHBoxLayout()
        search_row.addWidget(QLabel("🔍 Поиск:"))
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Начните вводить название продукта...")
        self.search_edit.setClearButtonEnabled(True)
        self.search_edit.textChanged.connect(self._on_search_changed)
        search_row.addWidget(self.search_edit)
        lay.addLayout(search_row)

        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels([
            "Продукт", "Ед. изм.", "Показывать Шт", "Кол-во в 1 шт",
            "От кол-ва шт", "Хвостик ШК", "Хвостик СД"
        ])
        hdr = self.table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(6, QHeaderView.ResizeMode.ResizeToContents)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        # Увеличиваем высоту строк в 3 раза (дефолт ~30px → 90px)
        self.table.verticalHeader().setDefaultSectionSize(90)
        lay.addWidget(self.table)

        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setWidget(content)
        main_lay = QVBoxLayout(self)
        main_lay.setContentsMargins(0, 0, 0, 0)
        main_lay.addWidget(scroll)

        self._load_table()

    def _on_search_changed(self, text: str):
        self._search_text = text.strip().lower()
        self._search_timer.start()

    def _apply_filter(self):
        """Скрывает строки, не совпадающие с поисковым запросом."""
        q = self._search_text
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            name = item.text().lower() if item else ""
            self.table.setRowHidden(row, bool(q) and q not in name)

    def _load_table(self):
        # Читаем данные один раз
        products = data_store.get_ref("products") or []
        eligible = sorted(
            [p for p in products if p.get("unit", "").lower() != "шт"],
            key=lambda p: p["name"].lower()
        )

        self._updating = True
        self.table.setRowCount(len(eligible))

        for row, prod in enumerate(eligible):
            name = prod["name"]

            # Название
            item_name = QTableWidgetItem(name)
            item_name.setFlags(item_name.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(row, 0, item_name)

            # Единица
            item_unit = QTableWidgetItem(prod.get("unit", ""))
            item_unit.setFlags(item_unit.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(row, 1, item_unit)

            show_pcs = prod.get("showPcs", False)

            # Показывать Шт — переключатель в центре ячейки
            chk = ToggleSwitch()
            chk.setChecked(show_pcs)
            chk_widget = QWidget()
            chk_lay = QHBoxLayout(chk_widget)
            chk_lay.addWidget(chk)
            chk_lay.setAlignment(Qt.AlignmentFlag.AlignCenter)
            chk_lay.setContentsMargins(0, 0, 0, 0)
            self.table.setCellWidget(row, 2, chk_widget)

            # Кол-во в 1 шт
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
            spin_min.setToolTip("Ниже этого количества — 0 шт. От этого и выше — расчёт по «Кол-во в 1 шт» и хвостику.")
            self.table.setCellWidget(row, 4, spin_min)

            # Хвостик ШК: от остатка (в ед. количества) не меньше — добавляем 1 шт
            spin_tail_shk = QDoubleSpinBox()
            spin_tail_shk.setRange(0, 99999.0)
            spin_tail_shk.setDecimals(3)
            spin_tail_shk.setSingleStep(0.01)
            v_shk = prod.get("roundTailFromШК")
            spin_tail_shk.setValue(v_shk if v_shk is not None else (0 if prod.get("roundUpШК", True) else float(prod.get("pcsPerUnit", 1))))
            spin_tail_shk.setEnabled(show_pcs)
            spin_tail_shk.setToolTip("Добавлять 1 шт, если остаток ≥ этого значения (ШК). 0 = в большую; ≥ вес 1 шт = в меньшую.")
            self.table.setCellWidget(row, 5, spin_tail_shk)

            # Хвостик СД
            spin_tail_sd = QDoubleSpinBox()
            spin_tail_sd.setRange(0, 99999.0)
            spin_tail_sd.setDecimals(3)
            spin_tail_sd.setSingleStep(0.01)
            v_sd = prod.get("roundTailFromСД")
            spin_tail_sd.setValue(v_sd if v_sd is not None else (0 if prod.get("roundUpСД", True) else float(prod.get("pcsPerUnit", 1))))
            spin_tail_sd.setEnabled(show_pcs)
            spin_tail_sd.setToolTip("Добавлять 1 шт, если остаток ≥ этого значения (СД). 0 = в большую; ≥ вес 1 шт = в меньшую.")
            self.table.setCellWidget(row, 6, spin_tail_sd)

            # Подключаем сигналы после установки значений
            chk.stateChanged.connect(
                lambda state, n=name, s=spin, sm=spin_min, t1=spin_tail_shk, t2=spin_tail_sd: self._on_show_pcs(n, state, s, sm, t1, t2)
            )
            spin.valueChanged.connect(
                lambda val, n=name: self._on_pcs_per_unit(n, val)
            )
            spin_min.valueChanged.connect(
                lambda val, n=name: self._on_min_qty_for_pcs(n, val)
            )
            spin_tail_shk.valueChanged.connect(
                lambda val, n=name: self._on_tail_shk(n, val)
            )
            spin_tail_sd.valueChanged.connect(
                lambda val, n=name: self._on_tail_sd(n, val)
            )

        self._updating = False

    def _on_show_pcs(self, name: str, state: int, spin: QDoubleSpinBox, spin_min: QDoubleSpinBox,
                     spin_tail_shk: QDoubleSpinBox, spin_tail_sd: QDoubleSpinBox):
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

    def refresh(self):
        self._load_table()
        if self.search_edit.text():
            self._apply_filter()
