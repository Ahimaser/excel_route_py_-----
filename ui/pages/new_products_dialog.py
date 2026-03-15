"""
new_products_dialog.py — Диалог выбора для новых названий продуктов (после обработки файлов).

Пользователь для каждого названия выбирает:
- «Новый продукт» — добавить в справочник (привязку к отделу сделать в «Отделы и продукты»).
- «Дубликат: <название>» — записать как вариант существующего; при следующих запусках
  это написание будет автоматически подставляться как выбранное каноническое.
"""
from __future__ import annotations

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTableWidget, QTableWidgetItem, QComboBox, QHeaderView,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QShortcut, QKeySequence

from ui.widgets import hint_icon_button

ACTION_NEW = "new"
ACTION_ALIAS_PREFIX = "alias:"


def run_new_products_dialog(parent, items: list[dict], all_canonical: list[str] | None = None) -> list[dict]:
    """
    items: список {name, unit, similar: list[str]} — новые названия из файлов, similar — похожие канонические.
    all_canonical: все канонические названия справочника (для выбора «Дубликат»).
    Возвращает: [{name, unit, action: "new"|"alias", canonical?: str}].
    """
    if not items:
        return []
    dlg = NewProductsDialog(parent, items, all_canonical or [])
    if dlg.exec() != QDialog.DialogCode.Accepted:
        return []
    return dlg.get_decisions()


class NewProductsDialog(QDialog):
    def __init__(self, parent, items: list[dict], all_canonical: list[str] | None = None):
        super().__init__(parent)
        self.setWindowTitle("Новые названия продуктов")
        self.setMinimumSize(620, 380)
        self.resize(760, 460)
        self._items = items
        self._all_canonical = list(all_canonical or [])
        self._combos: list[QComboBox] = []
        self._build_ui()

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 16, 20, 16)
        lay.setSpacing(12)
        title_row = QHBoxLayout()
        title_row.addWidget(QLabel("Обнаружены названия, которых нет в справочнике. Выберите действие для каждого."))
        title_row.addWidget(hint_icon_button(
            self,
            "«Новый продукт» — добавить в справочник; «Дубликат» — записать как вариант существующего продукта.",
            "Инструкция — Новые названия\n\n"
            "1. В таблице — названия из загруженного файла, которых нет в справочнике.\n"
            "2. «Новый продукт» — добавить в справочник (привязку к отделу сделаете в «Отделы и продукты»).\n"
            "3. «Дубликат: [каноническое]» — записать как вариант выбранного продукта; при следующих запусках это написание подставится автоматически.\n"
            "4. После выбора нажмите «Сохранить» — данные обновятся, можно продолжить обработку.",
            "Инструкция",
        ))
        title_row.addStretch()
        lay.addLayout(title_row)
        self.table = QTableWidget(len(self._items), 2)
        self.table.setHorizontalHeaderLabels(["Название из файла", "Действие"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.table.setMinimumHeight(max(140, min(320, 44 * min(len(self._items), 10))))
        for row, it in enumerate(self._items):
            name_item = QTableWidgetItem(f"{it['name']} ({it.get('unit', '')})")
            self.table.setItem(row, 0, name_item)
            combo = QComboBox()
            combo.addItem("Новый продукт", ACTION_NEW)
            similar = it.get("similar") or []
            # Сначала похожие канонические, затем остальные (без повторов)
            seen = set(similar)
            ordered_canonical = similar + [c for c in self._all_canonical if c not in seen]
            if ordered_canonical:
                combo.insertSeparator(1)
                for c in ordered_canonical:
                    combo.addItem(f"Дубликат: {c}", ACTION_ALIAS_PREFIX + c)

            def _on_combo_changed(idx, r=row, item=it, cbox=combo):
                val = cbox.currentData()
                if val == ACTION_NEW:
                    txt = f"{item['name']} ({item.get('unit', '')})"
                elif isinstance(val, str) and val.startswith(ACTION_ALIAS_PREFIX):
                    canonical = val[len(ACTION_ALIAS_PREFIX):]
                    txt = f"{item['name']} (→ {canonical})"
                else:
                    txt = f"{item['name']} ({item.get('unit', '')})"
                self.table.item(r, 0).setText(txt)

            self._combos.append(combo)
            combo.currentIndexChanged.connect(_on_combo_changed)
            self.table.setCellWidget(row, 1, combo)
        lay.addWidget(self.table)
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        ok_btn = QPushButton("Применить")
        ok_btn.setObjectName("btnPrimary")
        ok_btn.setDefault(True)
        ok_btn.setAutoDefault(True)
        ok_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("Отмена")
        cancel_btn.setObjectName("btnSecondary")
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(ok_btn)
        btn_row.addWidget(cancel_btn)
        lay.addLayout(btn_row)
        QShortcut(QKeySequence(Qt.Key.Key_Return), self, self.accept)
        if self._combos:
            self._combos[0].setFocus()

    def get_decisions(self) -> list[dict]:
        decisions = []
        for i, it in enumerate(self._items):
            combo = self._combos[i]
            val = combo.currentData()
            if val == ACTION_NEW:
                decisions.append({"name": it["name"], "unit": it.get("unit", ""), "action": "new"})
            elif isinstance(val, str) and val.startswith(ACTION_ALIAS_PREFIX):
                canonical = val[len(ACTION_ALIAS_PREFIX):]
                decisions.append({"name": it["name"], "unit": it.get("unit", ""), "action": "alias", "canonical": canonical})
            else:
                decisions.append({"name": it["name"], "unit": it.get("unit", ""), "action": "new"})
        return decisions
