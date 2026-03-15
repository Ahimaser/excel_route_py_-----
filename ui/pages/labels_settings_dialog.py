"""
labels_settings_dialog.py — Окно настройки этикеток: отделы/подотделы и условия.

Логика: галочка «Создавать этикетки» (labelsEnabled) у отдела/подотдела — включать его при создании этикеток.
Этикетки создаются без шаблонов (3 столбца: № маршрута, Дом, Количество).
"""
from __future__ import annotations

from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTreeWidget,
    QTreeWidgetItem,
    QWidget,
    QFrame,
    QFormLayout,
    QComboBox,
    QDoubleSpinBox,
    QLineEdit,
    QGroupBox,
    QScrollArea,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QShortcut, QKeySequence

from core import data_store
from ui.widgets import hint_icon_button, ToggleSwitch


class LabelsSettingsDialog(QDialog):
    """Диалог: галочки по отделам и условия этикеток (чищенка, сыпучка)."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Настройки этикеток")
        self.setMinimumSize(480, 420)
        self.resize(560, 500)
        self.setWindowModality(Qt.WindowModality.ApplicationModal)
        self.setSizeGripEnabled(True)
        self._updating = False
        self._build_ui()
        self._refresh_tree()
        self.tree.currentItemChanged.connect(self._on_selection_changed)
        self._current_node_obj = None

    def _save_and_close(self) -> None:
        self.accept()

    def _build_ui(self):
        content = QWidget()
        lay = QVBoxLayout(content)
        lay.setContentsMargins(20, 16, 20, 16)
        lay.setSpacing(12)
        title_row = QHBoxLayout()
        title_row.addWidget(QLabel("Настройки этикеток"))
        title_row.addWidget(hint_icon_button(
            self,
            "Галочка «Создавать этикетки» включает отдел. «Условия этикеток» — режимы чищенка и сыпучка.",
            "Инструкция — Настройки этикеток\n\n"
            "1. Галочка «Создавать этикетки» у отдела/подотдела — включать его при создании этикеток.\n"
            "2. «Условия этикеток для отдела» — настройка режимов: чищенка (деление по весу на этикетку), сыпучка (два файла по порогу).\n"
            "3. Этикетки создаются в формате 3 столбца: № маршрута, Дом/строение, Количество.",
            "Инструкция",
        ))
        title_row.addStretch()
        lay.addLayout(title_row)

        card = QFrame()
        card.setObjectName("card")
        card_lay = QVBoxLayout(card)
        card_lay.addWidget(QLabel("Отделы / подотделы"))
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Название", "Создавать этикетки"])
        self.tree.setColumnWidth(0, 280)
        self.tree.setColumnWidth(1, 160)
        self.tree.setMinimumHeight(280)
        card_lay.addWidget(self.tree)
        self.btn_label_rules = QPushButton("Условия этикеток для отдела…")
        self.btn_label_rules.setObjectName("btnSecondary")
        self.btn_label_rules.setMinimumWidth(220)
        self.btn_label_rules.setEnabled(False)
        self.btn_label_rules.clicked.connect(self._on_label_rules)
        card_lay.addWidget(self.btn_label_rules)
        self.lbl_no_depts = QLabel("Нет отделов. Добавьте отделы в меню «Файл» → «Справочники» → «Отделы и продукты».")
        self.lbl_no_depts.setObjectName("hintLabel")
        self.lbl_no_depts.setWordWrap(True)
        card_lay.addWidget(self.lbl_no_depts)
        lay.addWidget(card)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_save = QPushButton("Сохранить")
        btn_save.setObjectName("btnPrimary")
        btn_save.setDefault(True)
        btn_save.setAutoDefault(True)
        btn_save.clicked.connect(self._save_and_close)
        btn_row.addWidget(btn_save)
        lay.addLayout(btn_row)
        QShortcut(QKeySequence(Qt.Key.Key_Return), self, self._save_and_close)

        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setWidget(content)
        main_lay = QVBoxLayout(self)
        main_lay.setContentsMargins(0, 0, 0, 0)
        main_lay.addWidget(scroll)

    def _refresh_tree(self):
        self._updating = True
        try:
            self.tree.clear()
            depts = data_store.get_ref("departments") or []
            has_depts = False
            for dept in sorted((d for d in depts if isinstance(d, dict) and d.get("name")), key=lambda d: (d.get("name") or "").lower()):
                has_depts = True
                dept_item = QTreeWidgetItem([dept.get("name", ""), ""])
                dept_item.setData(0, Qt.ItemDataRole.UserRole, ("dept", dept))
                dept_item.setFlags(dept_item.flags() & ~Qt.ItemFlag.ItemIsUserCheckable)
                for sub in sorted((s for s in dept.get("subdepts", []) if isinstance(s, dict) and s.get("name")), key=lambda s: (s.get("name") or "").lower()):
                    sub_item = QTreeWidgetItem([sub.get("name", ""), ""])
                    sub_item.setData(0, Qt.ItemDataRole.UserRole, ("subdept", sub))
                    sub_item.setFlags(sub_item.flags() & ~Qt.ItemFlag.ItemIsUserCheckable)
                    dept_item.addChild(sub_item)
                self.tree.addTopLevelItem(dept_item)
                self._attach_toggle(dept_item, dept)
                for i in range(dept_item.childCount()):
                    child = dept_item.child(i)
                    child_data = child.data(0, Qt.ItemDataRole.UserRole)
                    if child_data and isinstance(child_data, (tuple, list)) and len(child_data) == 2:
                        self._attach_toggle(child, child_data[1])
                dept_item.setExpanded(True)
            self.lbl_no_depts.setVisible(not has_depts)
        finally:
            self._updating = False

    def _attach_toggle(self, item: QTreeWidgetItem, obj: dict) -> None:
        """Создаёт ToggleSwitch в колонке 1 дерева для объекта dept/subdept."""
        toggle = ToggleSwitch()
        toggle.setChecked(obj.get("labelsEnabled", True))
        toggle.stateChanged.connect(lambda state, o=obj: self._on_labels_toggle(o, state))
        container = QWidget()
        lay = QHBoxLayout(container)
        lay.setContentsMargins(8, 2, 8, 2)
        lay.addWidget(toggle)
        lay.addStretch()
        self.tree.setItemWidget(item, 1, container)

    def _on_labels_toggle(self, obj: dict, state: int) -> None:
        if self._updating:
            return
        obj["labelsEnabled"] = (state == 2)
        depts = data_store.get_ref("departments")
        if depts:
            data_store.set_key("departments", depts)

    def _on_selection_changed(self, current: QTreeWidgetItem | None, _prev):
        if not current:
            self._current_node_obj = None
            self.btn_label_rules.setEnabled(False)
            return
        data = current.data(0, Qt.ItemDataRole.UserRole)
        if not data or not isinstance(data, (tuple, list)) or len(data) != 2:
            self._current_node_obj = None
            self.btn_label_rules.setEnabled(False)
            return
        _, obj = data
        if not isinstance(obj, dict) or not obj.get("key"):
            self._current_node_obj = None
            self.btn_label_rules.setEnabled(False)
            return
        self._current_node_obj = obj
        self.btn_label_rules.setEnabled(True)

    def _on_label_rules(self):
        if not self._current_node_obj:
            return
        LabelRulesDialog(self._current_node_obj, self).exec()
        depts = data_store.get_ref("departments")
        if depts:
            data_store.set_key("departments", depts)


class LabelRulesDialog(QDialog):
    """Настройка правил этикеток для отдела/подотдела: чищенка (деление по весу), сыпучка (два файла по порогу)."""

    def __init__(self, node_obj: dict, parent=None):
        super().__init__(parent)
        self.node_obj = node_obj
        self.setWindowTitle(f"Условия этикеток: {node_obj.get('name', '')}")
        self.setMinimumWidth(420)
        lay = QVBoxLayout(self)

        self.combo_mode = QComboBox()
        self.combo_mode.addItem("По умолчанию (без особых правил)", "default")
        self.combo_mode.addItem("Чищенка (деление по весу на этикетке)", "chistchenka")
        self.combo_mode.addItem("Сыпучка (два файла: до/после порога)", "sypuchka")
        mode = node_obj.get("labelPrintMode") or "default"
        if mode not in ("chistchenka", "sypuchka"):
            mode = "default"
        idx = self.combo_mode.findData(mode)
        if idx >= 0:
            self.combo_mode.setCurrentIndex(idx)
        self.combo_mode.currentIndexChanged.connect(self._on_mode_changed)
        form = QFormLayout()
        form.addRow("Режим:", self.combo_mode)
        lay.addLayout(form)

        rules = node_obj.get("labelRules") or {}
        ch = rules.get("chistchenka") or {}
        sy = rules.get("sypuchka") or {}

        self.gr_ch = QGroupBox("Чищенка: параметры деления")
        form_ch = QFormLayout(self.gr_ch)
        self.spin_max_kg = QDoubleSpinBox()
        self.spin_max_kg.setRange(0.1, 100)
        self.spin_max_kg.setValue(float(ch.get("maxKgPerLabel", 5)))
        self.spin_max_kg.setSuffix(" кг")
        self.spin_max_kg.setToolTip(
            "Максимальный вес на одной этикетке. При весе 20 кг будет 4 этикетки по 5 кг; "
            "при 23 кг — 4 по 5 кг и последняя 3 кг. Каждый продукт — в отдельном файле: «продукт_дата_основной/увеличение»."
        )
        form_ch.addRow("Макс. вес на одну этикетку:", self.spin_max_kg)
        lay.addWidget(self.gr_ch)

        self.gr_sy = QGroupBox("Сыпучка: два файла по порогу")
        form_sy = QFormLayout(self.gr_sy)
        self.spin_threshold = QDoubleSpinBox()
        self.spin_threshold.setRange(0.1, 100)
        self.spin_threshold.setValue(float(sy.get("thresholdKg", 4)))
        self.spin_threshold.setSuffix(" кг")
        form_sy.addRow("Порог (кг):", self.spin_threshold)
        self.le_label_below = QLineEdit()
        self.le_label_below.setPlaceholderText("меньше 4 кг")
        self.le_label_below.setText(sy.get("labelBelow", "меньше 4 кг"))
        form_sy.addRow("Подпись для «≤ порога»:", self.le_label_below)
        self.le_label_above = QLineEdit()
        self.le_label_above.setPlaceholderText("больше 4 кг")
        self.le_label_above.setText(sy.get("labelAbove", "больше 4 кг"))
        form_sy.addRow("Подпись для «> порога»:", self.le_label_above)
        lay.addWidget(self.gr_sy)

        self._on_mode_changed()

        btn_lay = QHBoxLayout()
        btn_lay.addStretch()
        btn_ok = QPushButton("Сохранить")
        btn_ok.setObjectName("btnPrimary")
        btn_ok.setDefault(True)
        btn_ok.setAutoDefault(True)
        btn_ok.clicked.connect(self._save)
        btn_lay.addWidget(btn_ok)
        lay.addLayout(btn_lay)
        QShortcut(QKeySequence(Qt.Key.Key_Return), self, self._save)

    def _on_mode_changed(self):
        mode = self.combo_mode.currentData()
        self.gr_ch.setEnabled(mode == "chistchenka")
        self.gr_sy.setEnabled(mode == "sypuchka")
        self.spin_max_kg.setEnabled(mode == "chistchenka")
        for w in (self.spin_threshold, self.le_label_below, self.le_label_above):
            w.setEnabled(mode == "sypuchka")

    def _save(self):
        mode = self.combo_mode.currentData()
        self.node_obj["labelPrintMode"] = mode
        self.node_obj["labelRules"] = {
            "chistchenka": {
                "maxKgPerLabel": self.spin_max_kg.value(),
            },
            "sypuchka": {
                "thresholdKg": self.spin_threshold.value(),
                "labelBelow": (self.le_label_below.text() or "меньше 4 кг").strip(),
                "labelAbove": (self.le_label_above.text() or "больше 4 кг").strip(),
            },
        }
        self.accept()


def open_labels_settings_dialog(parent: QWidget):
    """Открывает модальное окно настроек этикеток."""
    dlg = LabelsSettingsDialog(parent)
    dlg.exec()
