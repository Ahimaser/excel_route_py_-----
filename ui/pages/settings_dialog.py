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
    QTableWidget, QTableWidgetItem, QCheckBox, QDoubleSpinBox,
    QComboBox, QHeaderView, QAbstractItemView, QLineEdit, QWidget
)
from PyQt6.QtCore import Qt, QTimer

from core import data_store
from ui.styles import STYLESHEET


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

        self.setStyleSheet(STYLESHEET)
        self._build_ui()
        QTimer.singleShot(0, self._load_table)

    # ─────────────────────────── Построение UI ────────────────────────────

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(14)

        # Заголовок
        lbl = QLabel("Настройки отображения в штуках")
        lbl.setObjectName("sectionTitle")
        lay.addWidget(lbl)

        lbl_hint = QLabel(
            "Настройки применяются ко всем продуктам с одинаковым названием. "
            "Отображение Шт доступно только для продуктов с единицей измерения, "
            "отличной от «шт»."
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
        self.search_edit.setToolTip("Фильтр по названию продукта. Таблица обновляется при вводе.")
        self.search_edit.textChanged.connect(self._on_search_changed)
        search_row.addWidget(self.search_edit)
        lay.addLayout(search_row)

        # Таблица
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels([
            "Продукт", "Ед. изм.", "Показывать Шт", "Кол-во в 1 шт",
            "Округление ШК", "Округление СД"
        ])
        hdr = self.table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
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
        btn_cancel.setToolTip("Закрыть окно без дополнительных действий")
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_cancel)

        btn_save = QPushButton("Сохранить и закрыть")
        btn_save.setObjectName("btnPrimary")
        btn_save.setFixedHeight(36)
        btn_save.setToolTip("Сохранить настройки и закрыть окно")
        btn_save.clicked.connect(self.accept)
        btn_row.addWidget(btn_save)

        lay.addLayout(btn_row)

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
        products = data_store.get("products") or []
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

            # Чекбокс «Показывать Шт»
            chk = QCheckBox()
            chk.setChecked(show_pcs)
            chk.setObjectName("tableCheckBox")
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

            # Округление ШК
            round_shk = prod.get("roundUpШК") if "roundUpШК" in prod else prod.get("roundUp", True)
            combo_shk = QComboBox()
            combo_shk.addItem("В большую сторону", True)
            combo_shk.addItem("В меньшую сторону", False)
            combo_shk.setCurrentIndex(0 if round_shk else 1)
            combo_shk.setEnabled(show_pcs)
            combo_shk.setToolTip("Округление штук для маршрутов ШК (школы)")
            self.table.setCellWidget(row, 4, combo_shk)

            # Округление СД
            round_sd = prod.get("roundUpСД") if "roundUpСД" in prod else prod.get("roundUp", True)
            combo_sd = QComboBox()
            combo_sd.addItem("В большую сторону", True)
            combo_sd.addItem("В меньшую сторону", False)
            combo_sd.setCurrentIndex(0 if round_sd else 1)
            combo_sd.setEnabled(show_pcs)
            combo_sd.setToolTip("Округление штук для маршрутов СД (сады)")
            self.table.setCellWidget(row, 5, combo_sd)

            # Сигналы
            chk.stateChanged.connect(
                lambda state, n=name, s=spin, c1=combo_shk, c2=combo_sd: self._on_show_pcs(n, state, s, c1, c2)
            )
            spin.valueChanged.connect(lambda val, n=name: self._on_pcs_per_unit(n, val))
            combo_shk.currentIndexChanged.connect(
                lambda idx, n=name, c=combo_shk: self._on_round_shk(n, c.currentData())
            )
            combo_sd.currentIndexChanged.connect(
                lambda idx, n=name, c=combo_sd: self._on_round_sd(n, c.currentData())
            )

        self._updating = False

    # ─────────────────────────── Обработчики изменений ────────────────────

    def _on_show_pcs(self, name: str, state: int, spin: QDoubleSpinBox,
                     combo_shk: QComboBox, combo_sd: QComboBox):
        if self._updating:
            return
        show = state == Qt.CheckState.Checked.value
        spin.setEnabled(show)
        combo_shk.setEnabled(show)
        combo_sd.setEnabled(show)
        data_store.update_product(name, showPcs=show)

    def _on_pcs_per_unit(self, name: str, val: float):
        if self._updating:
            return
        data_store.update_product(name, pcsPerUnit=val)

    def _on_round_shk(self, name: str, round_up: bool):
        if self._updating:
            return
        data_store.update_product(name, roundUpШК=round_up)

    def _on_round_sd(self, name: str, round_up: bool):
        if self._updating:
            return
        data_store.update_product(name, roundUpСД=round_up)


# ─────────────────────────── Публичная функция ────────────────────────────

def open_settings_dialog(parent, on_saved=None):
    """
    Открывает модальное окно настроек Шт.

    Args:
        parent:   родительский виджет (основное окно)
        on_saved: коллбэк без аргументов, вызывается при нажатии «Сохранить и закрыть»
    """
    dlg = SettingsDialog(parent)
    result = dlg.exec()
    if result == QDialog.DialogCode.Accepted and on_saved is not None:
        on_saved()
