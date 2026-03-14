"""
labels_settings_dialog.py — Окно настройки этикеток: отделы/подотделы и шаблоны по продуктам.

Логика работы с этикетками:
- Шаблоны этикеток задаются только здесь: для выбранного отдела/подотдела в таблице справа
  у каждого продукта можно указать файл шаблона XLS (или снять шаблон).
- Для каких отделов создавать этикетки — решает пользователь: в дереве слева у каждого
  отдела и подотдела есть галочка «Создавать этикетки» (labelsEnabled). Если галочка снята,
  продукты этого отдела/подотдела не попадают в генерацию этикеток, даже при наличии шаблона.
- При создании этикеток (страница «Этикетки») в файлы попадают только продукты, у которых:
  1) задан шаблон (labelTemplatePath), 2) отдел/подотдел продукта имеет labelsEnabled=True.
"""
from __future__ import annotations

import os
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTreeWidget, QTreeWidgetItem, QWidget, QTableWidget, QTableWidgetItem,
    QHeaderView, QAbstractItemView, QFileDialog, QFrame,
    QSplitter, QFormLayout, QComboBox, QDoubleSpinBox, QLineEdit, QGroupBox,
    QScrollArea, QStyle,
)
from PyQt6.QtCore import Qt, QSize

from core import data_store
from ui.widgets import hint_icon_button, ToggleSwitch
from ui.pages.label_template_editor import open_label_template_editor


class LabelsSettingsDialog(QDialog):
    """Диалог: печатать этикетки по отделам и шаблоны XLS по продуктам."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Настройки этикеток")
        self.setMinimumSize(900, 620)
        self.resize(1100, 720)
        self.setWindowModality(Qt.WindowModality.ApplicationModal)
        self.setSizeGripEnabled(True)
        self._updating = False
        self._build_ui()
        self._refresh_tree()
        self.tree.currentItemChanged.connect(self._on_selection_changed)
        self._current_node_obj = None

    def _build_ui(self):
        content = QWidget()
        lay = QVBoxLayout(content)
        lay.setContentsMargins(20, 16, 20, 16)
        lay.setSpacing(12)
        title_row = QHBoxLayout()
        title_row.addWidget(QLabel("Настройки этикеток"))
        title_row.addWidget(hint_icon_button(
            self,
            "Слева — галочка по отделам; справа — шаблон XLS для каждого продукта.",
            "Инструкция — Настройки этикеток\n\n"
            "1. Слева: дерево отделов и подотделов. Галочка «Создавать этикетки» — включать этот отдел при создании этикеток.\n"
            "2. Справа: выберите отдел или подотдел слева — отобразится таблица продуктов.\n"
            "3. Для каждого продукта укажите шаблон XLS: кнопка «…» — выбрать файл, «✕» — снять шаблон.\n"
            "4. Этикетки создаются только для продуктов с заданным шаблоном и с включённой галочкой у отдела.",
            "Инструкция",
        ))
        title_row.addStretch()
        lay.addLayout(title_row)

        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        left = QWidget()
        left_lay = QVBoxLayout(left)
        left_lay.setContentsMargins(0, 0, 0, 0)
        left_lay.addWidget(QLabel("Отделы / подотделы"))
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Название", "Создавать этикетки"])
        self.tree.setColumnWidth(0, 280)
        self.tree.setColumnWidth(1, 160)
        self.tree.setMinimumHeight(320)
        left_lay.addWidget(self.tree)
        self.btn_label_rules = QPushButton("Условия этикеток для отдела…")
        self.btn_label_rules.setObjectName("btnSecondary")
        self.btn_label_rules.setMinimumWidth(220)
        self.btn_label_rules.setEnabled(False)
        self.btn_label_rules.clicked.connect(self._on_label_rules)
        left_lay.addWidget(self.btn_label_rules)
        self.lbl_no_depts = QLabel("Нет отделов. Добавьте отделы в меню «Справочники» → «Отделы и продукты».")
        self.lbl_no_depts.setObjectName("hintLabel")
        self.lbl_no_depts.setWordWrap(True)
        left_lay.addWidget(self.lbl_no_depts)
        self.splitter.addWidget(left)

        right = QFrame()
        right.setObjectName("card")
        right_lay = QVBoxLayout(right)
        self.lbl_right_title = QLabel("Выберите отдел или подотдел слева")
        right_lay.addWidget(self.lbl_right_title)
        self.lbl_no_prods = QLabel("В выбранном отделе/подотделе нет продуктов.")
        self.lbl_no_prods.setObjectName("hintLabel")
        self.lbl_no_prods.setVisible(False)
        right_lay.addWidget(self.lbl_no_prods)
        self.products_table = QTableWidget()
        self.products_table.setColumnCount(3)
        self.products_table.setHorizontalHeaderLabels(["Продукт", "Шаблон этикетки (XLS)", ""])
        self.products_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.products_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.products_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        self.products_table.setColumnWidth(2, 130)
        self.products_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.products_table.verticalHeader().setVisible(False)
        self.products_table.setVisible(False)
        self.products_table.setMinimumHeight(320)
        right_lay.addWidget(self.products_table)
        self.splitter.addWidget(right)
        self.splitter.setSizes([420, 620])
        self.splitter.setStretchFactor(0, 1)
        self.splitter.setStretchFactor(1, 2)
        self.splitter.setChildrenCollapsible(False)
        lay.addWidget(self.splitter)

        btn_close = QPushButton("Закрыть")
        btn_close.setObjectName("btnSecondary")
        btn_close.clicked.connect(self.accept)
        lay.addWidget(btn_close, alignment=Qt.AlignmentFlag.AlignRight)

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
            self.lbl_right_title.setText("Выберите отдел или подотдел слева")
            self.products_table.setVisible(False)
            self._current_node_obj = None
            if hasattr(self, "btn_label_rules"):
                self.btn_label_rules.setEnabled(False)
            if hasattr(self, "lbl_no_prods"):
                self.lbl_no_prods.setVisible(False)
            return
        data = current.data(0, Qt.ItemDataRole.UserRole)
        if not data or not isinstance(data, (tuple, list)) or len(data) != 2:
            self.products_table.setVisible(False)
            self._current_node_obj = None
            if hasattr(self, "btn_label_rules"):
                self.btn_label_rules.setEnabled(False)
            if hasattr(self, "lbl_no_prods"):
                self.lbl_no_prods.setVisible(False)
            return
        _, obj = data
        if not isinstance(obj, dict):
            self.products_table.setVisible(False)
            self._current_node_obj = None
            if hasattr(self, "btn_label_rules"):
                self.btn_label_rules.setEnabled(False)
            if hasattr(self, "lbl_no_prods"):
                self.lbl_no_prods.setVisible(False)
            return
        key = obj.get("key")
        if not key:
            self.products_table.setVisible(False)
            self._current_node_obj = None
            if hasattr(self, "btn_label_rules"):
                self.btn_label_rules.setEnabled(False)
            if hasattr(self, "lbl_no_prods"):
                self.lbl_no_prods.setVisible(False)
            return
        self._current_node_obj = obj
        if hasattr(self, "btn_label_rules"):
            self.btn_label_rules.setEnabled(True)
        self.lbl_right_title.setText(f"Продукты: {obj.get('name', '')}")
        products = sorted(
            [p for p in (data_store.get_ref("products") or []) if p.get("deptKey") == key],
            key=lambda p: (p.get("name") or "").lower()
        )
        self._updating = True
        self.products_table.setRowCount(len(products))
        if hasattr(self, "lbl_no_prods"):
            self.lbl_no_prods.setVisible(len(products) == 0)
        for row, prod in enumerate(products):
            pname = prod.get("name", "")
            self.products_table.setItem(row, 0, QTableWidgetItem(f"{pname} ({prod.get('unit', '')})"))
            tpl = prod.get("labelTemplatePath") or ""
            lbl = QLabel(os.path.basename(tpl) if tpl else "—")
            btn = QPushButton()
            btn.setObjectName("btnIcon")
            btn.setFixedWidth(36)
            btn.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogOpenButton))
            btn.setIconSize(QSize(18, 18))
            btn.setToolTip("Выбрать файл шаблона XLS")
            btn_clear = QPushButton()
            btn_clear.setFixedWidth(36)
            btn_clear.setObjectName("btnIconDanger")
            trash_icon = getattr(QStyle.StandardPixmap, "SP_TrashIcon", QStyle.StandardPixmap.SP_DialogDiscardButton)
            btn_clear.setIcon(self.style().standardIcon(trash_icon))
            btn_clear.setIconSize(QSize(18, 18))
            btn_clear.setToolTip("Снять шаблон")
            btn_clear.setVisible(bool(tpl))
            btn.clicked.connect(lambda checked=False, n=pname, l=lbl, b=btn_clear: self._on_select_template(n, l, b))
            btn_clear.clicked.connect(lambda checked=False, n=pname, l=lbl, b=btn_clear: self._on_clear_template(n, l, b))
            cell = QWidget()
            cell_lay = QHBoxLayout(cell)
            cell_lay.setContentsMargins(4, 0, 4, 0)
            cell_lay.setSpacing(4)
            cell_lay.addWidget(lbl, 1)
            cell_lay.addWidget(btn)
            cell_lay.addWidget(btn_clear)
            self.products_table.setCellWidget(row, 1, cell)
            btn_preview = QPushButton("Предпросмотр")
            btn_preview.setObjectName("btnSecondary")
            btn_preview.setEnabled(bool(tpl))
            btn_preview.clicked.connect(lambda checked=False, n=pname: self._on_preview_template(n))
            self.products_table.setCellWidget(row, 2, btn_preview)
        self._updating = False
        self.products_table.setVisible(True)
        if hasattr(self, "lbl_no_prods"):
            self.lbl_no_prods.setVisible(len(products) == 0)

    def _on_select_template(self, product_name: str, label_widget: QLabel, btn_clear: QPushButton | None = None):
        start_dir = ""
        products = data_store.get_ref("products") or []
        prod = next((p for p in products if p.get("name") == product_name), None)
        if prod and prod.get("labelTemplatePath") and os.path.isfile(prod["labelTemplatePath"]):
            start_dir = os.path.dirname(prod["labelTemplatePath"])
        path, _ = QFileDialog.getOpenFileName(
            self, "Шаблон этикетки (XLS)",
            start_dir or None,
            "Excel 97-2003 (*.xls);;Все файлы (*)",
        )
        if path:
            path = os.path.normpath(path)
            if path.lower().endswith(".xlsx"):
                QMessageBox.warning(
                    self, "Формат файла",
                    "Поддерживается только формат Excel 97-2003 (.xls). Выберите файл .xls или сохраните шаблон в этом формате.",
                )
                return
            if data_store.update_product(product_name, labelTemplatePath=path):
                label_widget.setText(os.path.basename(path))
                if btn_clear is not None:
                    btn_clear.setVisible(True)
                for row in range(self.products_table.rowCount()):
                    it = self.products_table.item(row, 0)
                    if it and product_name in it.text():
                        w = self.products_table.cellWidget(row, 2)
                        if isinstance(w, QPushButton):
                            w.setEnabled(True)
                        break
            else:
                QMessageBox.warning(
                    self, "Ошибка",
                    f"Не удалось привязать шаблон к продукту «{product_name}». Проверьте, что продукт есть в справочнике.",
                )

    def _on_clear_template(self, product_name: str, label_widget: QLabel, btn_clear: QPushButton):
        data_store.update_product(product_name, labelTemplatePath="")
        label_widget.setText("—")
        btn_clear.setVisible(False)
        for row in range(self.products_table.rowCount()):
            it = self.products_table.item(row, 0)
            if it and product_name in it.text():
                w = self.products_table.cellWidget(row, 2)
                if isinstance(w, QPushButton):
                    w.setEnabled(False)
                break

    def _on_preview_template(self, product_name: str):
        open_label_template_editor(product_name, self)

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
        btn_ok.clicked.connect(self._save)
        btn_lay.addWidget(btn_ok)
        lay.addLayout(btn_lay)

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
