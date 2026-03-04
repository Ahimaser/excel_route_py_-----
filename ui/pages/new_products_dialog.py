"""
new_products_dialog.py — Диалог выбора действия для новых названий продуктов из файлов.
"""
from __future__ import annotations

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTableWidget, QTableWidgetItem, QComboBox, QHeaderView,
)
from PyQt6.QtCore import Qt

from ui.styles import STYLESHEET

# Ключи выбора в комбо
ACTION_NEW = "new"
ACTION_COPY = "copy"
ACTION_ALIAS_PREFIX = "alias:"


def run_new_products_dialog(parent, items: list[dict]) -> list[dict]:
    """
    items: список {name, unit, similar: list[str]} — похожие канонические названия.
    Возвращает список решений: {name, unit, action: "new"|"copy"|"alias", canonical?: str}.
    """
    if not items:
        return []
    dlg = NewProductsDialog(parent, items)
    if dlg.exec() != QDialog.DialogCode.Accepted:
        return []
    return dlg.get_decisions()


class NewProductsDialog(QDialog):
    def __init__(self, parent, items: list[dict]):
        super().__init__(parent)
        self.setWindowTitle("Новые названия продуктов")
        self.setMinimumSize(560, 340)
        self.resize(640, 400)
        self._items = items
        self._combos: list[QComboBox] = []
        self.setStyleSheet(STYLESHEET)
        self._build_ui()

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 16, 20, 16)
        lay.setSpacing(12)
        hint = QLabel(
            "Обнаружены названия, которых нет в справочнике. "
            "Добавьте как вариант к существующему продукту или создайте новый. "
            "Копию можно потом перетащить к существующему в окне «Продукты»."
        )
        hint.setObjectName("stepLabel")
        hint.setWordWrap(True)
        lay.addWidget(hint)
        self.table = QTableWidget(len(self._items), 2)
        self.table.setHorizontalHeaderLabels(["Название из файла", "Действие"])
        self.table.setToolTip("Для каждого нового названия выберите: добавить как новый продукт, копию или вариант существующего")
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.table.setMinimumHeight(max(120, min(280, 44 * min(len(self._items), 8))))
        for row, it in enumerate(self._items):
            self.table.setItem(row, 0, QTableWidgetItem(f"{it['name']}  ({it.get('unit', '')})"))
            combo = QComboBox()
            combo.addItem("Новый продукт", ACTION_NEW)
            combo.addItem("Копия (связать вручную в окне Продукты)", ACTION_COPY)
            similar = it.get("similar") or []
            if similar:
                combo.insertSeparator(2)
                for c in similar:
                    combo.addItem(f"Добавить к: {c}", ACTION_ALIAS_PREFIX + c)
            # по умолчанию — «Новый продукт» (индекс 0)
            self._combos.append(combo)
            self.table.setCellWidget(row, 1, combo)
        lay.addWidget(self.table)
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        ok_btn = QPushButton("Применить")
        ok_btn.setObjectName("btnPrimary")
        ok_btn.setToolTip("Сохранить выбранные действия для всех новых названий")
        ok_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("Отмена")
        cancel_btn.setObjectName("btnSecondary")
        cancel_btn.setToolTip("Отменить и не добавлять новые продукты в справочник")
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(ok_btn)
        btn_row.addWidget(cancel_btn)
        lay.addLayout(btn_row)
        if self._combos:
            self._combos[0].setFocus()

    def get_decisions(self) -> list[dict]:
        decisions = []
        for i, it in enumerate(self._items):
            combo = self._combos[i]
            val = combo.currentData()
            if val == ACTION_NEW:
                decisions.append({"name": it["name"], "unit": it.get("unit", ""), "action": "new"})
            elif val == ACTION_COPY:
                decisions.append({"name": it["name"], "unit": it.get("unit", ""), "action": "copy"})
            elif isinstance(val, str) and val.startswith(ACTION_ALIAS_PREFIX):
                canonical = val[len(ACTION_ALIAS_PREFIX):]
                decisions.append({"name": it["name"], "unit": it.get("unit", ""), "action": "alias", "canonical": canonical})
            else:
                decisions.append({"name": it["name"], "unit": it.get("unit", ""), "action": "new"})
        return decisions
