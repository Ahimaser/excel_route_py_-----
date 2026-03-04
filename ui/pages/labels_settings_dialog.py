"""
labels_settings_dialog.py — Окно настройки этикеток: отделы/подотделы и шаблоны по продуктам.
"""
from __future__ import annotations

import os
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTreeWidget, QTreeWidgetItem, QWidget, QTableWidget, QTableWidgetItem,
    QCheckBox, QHeaderView, QAbstractItemView, QFileDialog, QFrame,
)
from PyQt6.QtCore import Qt

from core import data_store
from ui.styles import STYLESHEET


class LabelsSettingsDialog(QDialog):
    """Диалог: печатать этикетки по отделам и шаблоны XLS по продуктам."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Настройки этикеток")
        self.setMinimumSize(720, 500)
        self.resize(860, 560)
        self.setWindowModality(Qt.WindowModality.ApplicationModal)
        self._updating = False
        self.setStyleSheet(STYLESHEET)
        self._build_ui()
        self._refresh_tree()
        self.tree.currentItemChanged.connect(self._on_selection_changed)

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 16, 20, 16)
        lay.setSpacing(12)
        lay.addWidget(QLabel("Отделы/подотделы — печать этикеток и шаблоны по продуктам"))
        splitter = QHBoxLayout()
        left = QWidget()
        left_lay = QVBoxLayout(left)
        left_lay.addWidget(QLabel("Отделы / подотделы"))
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Название", "Печатать этикетки"])
        self.tree.setColumnWidth(0, 200)
        self.tree.setColumnWidth(1, 130)
        left_lay.addWidget(self.tree)
        self.lbl_no_depts = QLabel("Нет отделов. Добавьте отделы в меню «Справочники» → «Отделы и продукты».")
        self.lbl_no_depts.setObjectName("hintLabel")
        self.lbl_no_depts.setWordWrap(True)
        left_lay.addWidget(self.lbl_no_depts)
        splitter.addWidget(left, 1)

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
        self.products_table.setColumnCount(2)
        self.products_table.setHorizontalHeaderLabels(["Продукт", "Шаблон этикетки (XLS)"])
        self.products_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.products_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.products_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.products_table.verticalHeader().setVisible(False)
        self.products_table.setVisible(False)
        right_lay.addWidget(self.products_table)
        splitter.addWidget(right, 2)
        lay.addLayout(splitter)

        btn_close = QPushButton("Закрыть")
        btn_close.setObjectName("btnSecondary")
        btn_close.clicked.connect(self.accept)
        lay.addWidget(btn_close, alignment=Qt.AlignmentFlag.AlignRight)

    def _refresh_tree(self):
        self.tree.clear()
        depts = data_store.get_ref("departments") or []
        has_depts = False
        for dept in sorted((d for d in depts if isinstance(d, dict) and d.get("name")), key=lambda d: (d.get("name") or "").lower()):
            has_depts = True
            dept_item = QTreeWidgetItem([dept.get("name", ""), ""])
            dept_item.setData(0, Qt.ItemDataRole.UserRole, ("dept", dept))
            chk = QCheckBox()
            chk.setChecked(dept.get("labelsEnabled", True))
            chk.stateChanged.connect(lambda s, o=dept: self._on_labels_enabled(o, s == Qt.CheckState.Checked.value))
            w = QWidget()
            QHBoxLayout(w).addWidget(chk)
            self.tree.setItemWidget(dept_item, 1, w)
            for sub in sorted((s for s in dept.get("subdepts", []) if isinstance(s, dict) and s.get("name")), key=lambda s: (s.get("name") or "").lower()):
                sub_item = QTreeWidgetItem([f"  {sub.get('name', '')}", ""])
                sub_item.setData(0, Qt.ItemDataRole.UserRole, ("subdept", sub))
                chk2 = QCheckBox()
                chk2.setChecked(sub.get("labelsEnabled", True))
                chk2.stateChanged.connect(lambda s, o=sub: self._on_labels_enabled(o, s == Qt.CheckState.Checked.value))
                w2 = QWidget()
                QHBoxLayout(w2).addWidget(chk2)
                self.tree.setItemWidget(sub_item, 1, w2)
                dept_item.addChild(sub_item)
            self.tree.addTopLevelItem(dept_item)
            dept_item.setExpanded(True)
        self.lbl_no_depts.setVisible(not has_depts)

    def _on_labels_enabled(self, obj: dict, enabled: bool):
        if self._updating:
            return
        key = obj.get("key")
        if not key:
            return
        depts = data_store.get("departments") or []
        found = False
        for d in depts:
            if d.get("key") == key:
                d["labelsEnabled"] = enabled
                found = True
                break
            for s in d.get("subdepts", []):
                if s.get("key") == key:
                    s["labelsEnabled"] = enabled
                    found = True
                    break
            if found:
                break
        if found:
            data_store.set_key("departments", depts)

    def _on_selection_changed(self, current: QTreeWidgetItem | None, _prev):
        if not current:
            self.lbl_right_title.setText("Выберите отдел или подотдел слева")
            self.products_table.setVisible(False)
            if hasattr(self, "lbl_no_prods"):
                self.lbl_no_prods.setVisible(False)
            return
        data = current.data(0, Qt.ItemDataRole.UserRole)
        if not data or not isinstance(data, (tuple, list)) or len(data) != 2:
            self.products_table.setVisible(False)
            if hasattr(self, "lbl_no_prods"):
                self.lbl_no_prods.setVisible(False)
            return
        _, obj = data
        if not isinstance(obj, dict):
            self.products_table.setVisible(False)
            if hasattr(self, "lbl_no_prods"):
                self.lbl_no_prods.setVisible(False)
            return
        key = obj.get("key")
        if not key:
            self.products_table.setVisible(False)
            if hasattr(self, "lbl_no_prods"):
                self.lbl_no_prods.setVisible(False)
            return
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
            lbl.setToolTip(tpl or "Шаблон не выбран")
            btn = QPushButton("…")
            btn.setFixedWidth(32)
            btn.setToolTip("Выбрать шаблон XLS")
            btn.clicked.connect(lambda checked=False, n=pname, l=lbl: self._on_select_template(n, l))
            btn_clear = QPushButton("✕")
            btn_clear.setFixedWidth(28)
            btn_clear.setToolTip("Снять шаблон")
            btn_clear.setObjectName("btnIconDanger")
            btn_clear.setVisible(bool(tpl))
            btn_clear.clicked.connect(lambda checked=False, n=pname, l=lbl, b=btn_clear: self._on_clear_template(n, l, b))
            cell = QWidget()
            cell_lay = QHBoxLayout(cell)
            cell_lay.setContentsMargins(4, 0, 4, 0)
            cell_lay.setSpacing(4)
            cell_lay.addWidget(lbl, 1)
            cell_lay.addWidget(btn)
            cell_lay.addWidget(btn_clear)
            self.products_table.setCellWidget(row, 1, cell)
        self._updating = False
        self.products_table.setVisible(True)
        if hasattr(self, "lbl_no_prods"):
            self.lbl_no_prods.setVisible(len(products) == 0)

    def _on_select_template(self, product_name: str, label_widget: QLabel):
        start_dir = os.path.dirname(label_widget.toolTip()) if label_widget.toolTip() and os.path.isfile(label_widget.toolTip()) else ""
        path, _ = QFileDialog.getOpenFileName(self, "Шаблон этикетки (XLS)", start_dir, "Excel 97-2003 (*.xls)")
        if path:
            data_store.update_product(product_name, labelTemplatePath=path)
            label_widget.setText(os.path.basename(path))
            label_widget.setToolTip(path)

    def _on_clear_template(self, product_name: str, label_widget: QLabel, btn_clear: QPushButton):
        data_store.update_product(product_name, labelTemplatePath="")
        label_widget.setText("—")
        label_widget.setToolTip("Шаблон не выбран")
        btn_clear.setVisible(False)


def open_labels_settings_dialog(parent: QWidget):
    """Открывает модальное окно настроек этикеток."""
    dlg = LabelsSettingsDialog(parent)
    dlg.exec()
