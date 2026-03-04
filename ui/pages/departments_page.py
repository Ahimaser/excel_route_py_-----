"""
departments_page.py — Управление отделами, подотделами и продуктами.

Открывается как модальный диалог (open_modal).
Основное окно блокируется пока открыто это окно.
"""
from __future__ import annotations

from PyQt6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTreeWidget, QTreeWidgetItem,
    QLineEdit, QComboBox, QFormLayout, QDialogButtonBox,
    QMessageBox, QInputDialog, QListWidget, QListWidgetItem
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont

from core import data_store
from ui.styles import STYLESHEET
from ui.widgets import CommitLineEdit


class DepartmentsDialog(QDialog):
    """Модальный диалог управления отделами и продуктами."""

    def __init__(self, app_state: dict, parent=None):
        super().__init__(parent)
        self.app_state = app_state
        self.setWindowTitle("Отделы и продукты")
        self.setMinimumSize(780, 560)
        self.setModal(True)
        self.setStyleSheet(STYLESHEET)
        self._build_ui()
        self._refresh_tree()

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 16, 20, 16)
        lay.setSpacing(12)

        lbl = QLabel("Отделы и продукты")
        lbl.setObjectName("sectionTitle")
        lay.addWidget(lbl)

        hint = QLabel(
            "Здесь настраивается иерархия: Отдел → Подотдел → Продукт. "
            "Продукты привязываются к отделу или подотделу. "
            "✏ — переименовать; у продукта: ✕ — открепить от отдела, 🗑 — удалить из справочника; у отдела/подотдела: ✕ — удалить."
        )
        hint.setObjectName("hintLabel")
        hint.setWordWrap(True)
        lay.addWidget(hint)

        btn_bar = QHBoxLayout()
        btn_add_dept = QPushButton("+ Добавить отдел")
        btn_add_dept.setObjectName("btnPrimary")
        btn_add_dept.setToolTip("Создать новый отдел верхнего уровня")
        btn_add_dept.clicked.connect(self._add_dept)
        btn_bar.addWidget(btn_add_dept)

        btn_add_subdept = QPushButton("+ Добавить подотдел")
        btn_add_subdept.setObjectName("btnSecondary")
        btn_add_subdept.setToolTip("Создать подотдел внутри существующего отдела")
        btn_add_subdept.clicked.connect(self._add_subdept)
        btn_bar.addWidget(btn_add_subdept)

        btn_add_prod = QPushButton("+ Привязать продукт")
        btn_add_prod.setObjectName("btnSecondary")
        btn_add_prod.setToolTip(
            "Привязать один или несколько продуктов к отделу/подотделу.\n"
            "Используйте Ctrl+клик или Shift+клик для выбора нескольких."
        )
        btn_add_prod.clicked.connect(self._add_product)
        btn_bar.addWidget(btn_add_prod)

        btn_bar.addStretch()
        lay.addLayout(btn_bar)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Название", "Тип", "Продуктов", "Действия"])
        self.tree.setColumnWidth(0, 320)
        self.tree.setColumnWidth(1, 100)
        self.tree.setColumnWidth(2, 90)
        self.tree.setColumnWidth(3, 120)
        self.tree.setAlternatingRowColors(True)
        self.tree.setToolTip(
            "Дерево отделов, подотделов и продуктов.\n"
            "✏ — переименовать; ✕ — открепить/удалить; 🗑 — у продукта удалить из справочника."
        )
        lay.addWidget(self.tree)

        btn_close = QPushButton("Закрыть")
        btn_close.setObjectName("btnSecondary")
        btn_close.setToolTip("Закрыть окно и вернуться в основное приложение")
        btn_close.clicked.connect(self.accept)
        lay.addWidget(btn_close, alignment=Qt.AlignmentFlag.AlignRight)

    # ─────────────────────────── Дерево ───────────────────────────────────

    def _build_product_index(self, products: list) -> dict:
        index: dict[str, int] = {}
        for p in products:
            k = p.get("deptKey")
            if k:
                index[k] = index.get(k, 0) + 1
        return index

    def _refresh_tree(self):
        self.tree.setUpdatesEnabled(False)
        self.tree.clear()

        depts    = data_store.get_ref("departments") or []
        products = data_store.get_ref("products") or []
        prod_index = self._build_product_index(products)
        bold_font = QFont("", 13, QFont.Weight.Bold)

        for dept in sorted(depts, key=lambda d: d["name"].lower()):
            dept_item = QTreeWidgetItem([
                dept["name"], "Отдел",
                str(prod_index.get(dept["key"], 0)), ""
            ])
            dept_item.setData(0, Qt.ItemDataRole.UserRole, dept)
            dept_item.setFont(0, bold_font)
            dept_item.setToolTip(0, f"Отдел: {dept['name']}")
            self._add_action_buttons(dept_item, "dept", dept["key"])

            for sub in sorted(dept.get("subdepts", []), key=lambda s: s["name"].lower()):
                sub_item = QTreeWidgetItem([
                    f"  {sub['name']}", "Подотдел",
                    str(prod_index.get(sub["key"], 0)), ""
                ])
                sub_item.setData(0, Qt.ItemDataRole.UserRole, sub)
                sub_item.setToolTip(0, f"Подотдел: {sub['name']}")
                self._add_action_buttons(sub_item, "subdept", sub["key"])

                for prod in sorted(
                    (p for p in products if p.get("deptKey") == sub["key"]),
                    key=lambda p: p["name"].lower()
                ):
                    prod_item = QTreeWidgetItem([
                        f"    📦 {prod['name']}", "Продукт",
                        prod.get("unit", ""), ""
                    ])
                    prod_item.setData(0, Qt.ItemDataRole.UserRole, prod)
                    prod_item.setToolTip(0, f"Продукт: {prod['name']}, ед. изм.: {prod.get('unit','')}")
                    self._add_action_buttons(prod_item, "product", prod["name"])
                    sub_item.addChild(prod_item)

                dept_item.addChild(sub_item)

            for prod in sorted(
                (p for p in products if p.get("deptKey") == dept["key"]),
                key=lambda p: p["name"].lower()
            ):
                prod_item = QTreeWidgetItem([
                    f"  📦 {prod['name']}", "Продукт",
                    prod.get("unit", ""), ""
                ])
                prod_item.setData(0, Qt.ItemDataRole.UserRole, prod)
                prod_item.setToolTip(0, f"Продукт: {prod['name']}, ед. изм.: {prod.get('unit','')}")
                self._add_action_buttons(prod_item, "product", prod["name"])
                dept_item.addChild(prod_item)

            self.tree.addTopLevelItem(dept_item)
            dept_item.setExpanded(True)

        self.tree.setUpdatesEnabled(True)

    def _add_action_buttons(self, item: QTreeWidgetItem, item_type: str, key: str):
        btn_widget = QWidget()
        btn_lay = QHBoxLayout(btn_widget)
        btn_lay.setContentsMargins(4, 2, 4, 2)
        btn_lay.setSpacing(4)

        if item_type in ("dept", "subdept"):
            btn_rename = QPushButton("✏")
            btn_rename.setObjectName("btnIcon")
            btn_rename.setToolTip("Переименовать этот элемент")
            btn_rename.setFixedSize(28, 28)
            btn_rename.clicked.connect(lambda _, k=key, t=item_type: self._rename(k, t))
            btn_lay.addWidget(btn_rename)

        if item_type == "product":
            btn_unlink = QPushButton("✕")
            btn_unlink.setObjectName("btnIconDanger")
            btn_unlink.setToolTip("Открепить от отдела (продукт остаётся в справочнике)")
            btn_unlink.setFixedSize(28, 28)
            btn_unlink.clicked.connect(lambda _, k=key, t=item_type: self._delete(k, t))
            btn_lay.addWidget(btn_unlink)
            btn_remove = QPushButton("🗑")
            btn_remove.setObjectName("btnIconDanger")
            btn_remove.setToolTip("Удалить продукт из справочника полностью")
            btn_remove.setFixedSize(28, 28)
            btn_remove.clicked.connect(lambda _, k=key: self._delete_product_from_ref(k))
            btn_lay.addWidget(btn_remove)
        else:
            btn_del = QPushButton("✕")
            btn_del.setObjectName("btnIconDanger")
            btn_del.setToolTip("Удалить отдел/подотдел (продукты будут откреплены)")
            btn_del.setFixedSize(28, 28)
            btn_del.clicked.connect(lambda _, k=key, t=item_type: self._delete(k, t))
            btn_lay.addWidget(btn_del)

        btn_lay.addStretch()
        self.tree.setItemWidget(item, 3, btn_widget)

    # ─────────────────────────── Действия ─────────────────────────────────

    def _add_dept(self):
        name, ok = QInputDialog.getText(
            self, "Новый отдел",
            "Введите название нового отдела:"
        )
        if not ok or not name.strip():
            return
        name = name.strip()
        depts = data_store.get("departments") or []
        key = f"dept_{len(depts)}_{name[:10]}"
        depts.append({"key": key, "name": name, "subdepts": [], "labelsEnabled": True})
        data_store.set_key("departments", depts)
        self._refresh_tree()

    def _add_subdept(self):
        depts = data_store.get("departments") or []
        if not depts:
            QMessageBox.information(self, "Нет отделов",
                                    "Сначала создайте хотя бы один отдел.")
            return

        dlg = QDialog(self)
        dlg.setWindowTitle("Новый подотдел")
        dlg.setMinimumWidth(380)
        form = QFormLayout(dlg)

        combo_dept = QComboBox()
        combo_dept.setToolTip("Выберите отдел, в который добавить подотдел")
        for d in sorted(depts, key=lambda d: d["name"].lower()):
            combo_dept.addItem(d["name"], d["key"])
        form.addRow("Отдел:", combo_dept)

        le_name = CommitLineEdit()
        le_name.setToolTip("Введите название нового подотдела. Enter — подтвердить.")
        form.addRow("Название подотдела:", le_name)

        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        # Enter в поле названия также подтверждает диалог
        le_name.commit.connect(dlg.accept)
        btns.accepted.connect(dlg.accept)
        btns.rejected.connect(dlg.reject)
        form.addRow(btns)

        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        name = le_name.text().strip()
        if not name:
            return
        dept_key = combo_dept.currentData()

        for dept in depts:
            if dept["key"] == dept_key:
                subs = dept.get("subdepts", [])
                sub_key = f"sub_{dept_key}_{len(subs)}_{name[:10]}"
                sub_obj = {"key": sub_key, "name": name, "labelsEnabled": True}
                if "чищенка" in name.lower():
                    sub_obj["labelPrintMode"] = "chistchenka"
                elif "сыпучка" in name.lower():
                    sub_obj["labelPrintMode"] = "sypuchka"
                subs.append(sub_obj)
                dept["subdepts"] = subs
                break

        data_store.set_key("departments", depts)
        self._refresh_tree()

    def _add_product(self):
        depts    = data_store.get("departments") or []
        products = data_store.get("products") or []

        unassigned = sorted(
            [p for p in products if not p.get("deptKey")],
            key=lambda p: p["name"].lower()
        )
        if not unassigned:
            QMessageBox.information(
                self, "Нет свободных продуктов",
                "Все продукты уже привязаны к отделам.\n"
                "Загрузите XLS-файлы на главной странице, чтобы получить новые продукты."
            )
            return
        if not depts:
            QMessageBox.information(self, "Нет отделов",
                                    "Сначала создайте хотя бы один отдел.")
            return

        dlg = QDialog(self)
        dlg.setWindowTitle("Привязать продукты к отделу")
        dlg.setMinimumWidth(500)
        dlg.setMinimumHeight(440)
        vlay = QVBoxLayout(dlg)
        vlay.setSpacing(10)

        lbl_target = QLabel("Привязать к отделу / подотделу:")
        lbl_target.setToolTip("Выберите, к какому отделу или подотделу привязать продукты")
        vlay.addWidget(lbl_target)

        combo_target = QComboBox()
        combo_target.setToolTip("Отдел или подотдел для привязки продуктов")
        for d in sorted(depts, key=lambda d: d["name"].lower()):
            combo_target.addItem(f"• {d['name']}", d["key"])
            for sub in sorted(d.get("subdepts", []), key=lambda s: s["name"].lower()):
                combo_target.addItem(f"  └ {sub['name']}", sub["key"])
        vlay.addWidget(combo_target)

        lbl_prods = QLabel("Продукты (Ctrl+клик или Shift+клик — выбрать несколько):")
        lbl_prods.setToolTip("Выберите один или несколько продуктов для привязки")
        vlay.addWidget(lbl_prods)

        list_prods = QListWidget()
        list_prods.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        list_prods.setToolTip(
            "Список непривязанных продуктов.\n"
            "Ctrl+клик — выбрать несколько, Shift+клик — диапазон."
        )
        for p in unassigned:
            item = QListWidgetItem(f"{p['name']} ({p.get('unit', '')})")
            item.setData(Qt.ItemDataRole.UserRole, p["name"])
            list_prods.addItem(item)
        vlay.addWidget(list_prods)

        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        btns.accepted.connect(dlg.accept)
        btns.rejected.connect(dlg.reject)
        vlay.addWidget(btns)

        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        selected_names = {
            item.data(Qt.ItemDataRole.UserRole)
            for item in list_prods.selectedItems()
        }
        if not selected_names:
            return

        target_key = combo_target.currentData()
        for p in products:
            if p["name"] in selected_names:
                p["deptKey"] = target_key

        data_store.set_key("products", products)
        self._refresh_tree()

    def _rename(self, key: str, item_type: str):
        depts = data_store.get("departments") or []
        if item_type == "dept":
            for d in depts:
                if d["key"] == key:
                    name, ok = QInputDialog.getText(
                        self, "Переименовать отдел",
                        "Новое название отдела:", text=d["name"]
                    )
                    if ok and name.strip():
                        d["name"] = name.strip()
                    break
        else:
            for d in depts:
                for sub in d.get("subdepts", []):
                    if sub["key"] == key:
                        name, ok = QInputDialog.getText(
                            self, "Переименовать подотдел",
                            "Новое название подотдела:", text=sub["name"]
                        )
                        if ok and name.strip():
                            sub["name"] = name.strip()
                        break

        data_store.set_key("departments", depts)
        self._refresh_tree()

    def _delete(self, key: str, item_type: str):
        if item_type == "product":
            reply = QMessageBox.question(
                self, "Открепить продукт",
                "Открепить продукт от отдела?\n"
                "Продукт останется в справочнике и его можно будет привязать заново.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply != QMessageBox.StandardButton.Yes:
                return
            products = data_store.get("products") or []
            for p in products:
                if p["name"] == key:
                    p["deptKey"] = None
                    break
            data_store.set_key("products", products)

        elif item_type == "dept":
            reply = QMessageBox.question(
                self, "Удалить отдел",
                "Удалить отдел и все его подотделы?\n"
                "Все привязанные продукты будут откреплены.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply != QMessageBox.StandardButton.Yes:
                return
            depts    = data_store.get("departments") or []
            products = data_store.get("products") or []
            dept_to_del = next((d for d in depts if d["key"] == key), None)
            if dept_to_del:
                all_keys = {dept_to_del["key"]} | {s["key"] for s in dept_to_del.get("subdepts", [])}
                for p in products:
                    if p.get("deptKey") in all_keys:
                        p["deptKey"] = None
                depts = [d for d in depts if d["key"] != key]
                data_store.set_key("departments", depts)
                data_store.set_key("products", products)

        elif item_type == "subdept":
            reply = QMessageBox.question(
                self, "Удалить подотдел",
                "Удалить подотдел?\nВсе привязанные к нему продукты будут откреплены.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply != QMessageBox.StandardButton.Yes:
                return
            depts    = data_store.get("departments") or []
            products = data_store.get("products") or []
            for d in depts:
                subs = d.get("subdepts", [])
                sub = next((s for s in subs if s["key"] == key), None)
                if sub:
                    for p in products:
                        if p.get("deptKey") == key:
                            p["deptKey"] = None
                    d["subdepts"] = [s for s in subs if s["key"] != key]
                    break
            data_store.set_key("departments", depts)
            data_store.set_key("products", products)

        self._refresh_tree()

    def _delete_product_from_ref(self, product_name: str):
        """Удаляет продукт из справочника полностью (и связанные алиасы)."""
        reply = QMessageBox.question(
            self, "Удалить из справочника",
            f"Удалить продукт «{product_name}» из справочника полностью?\n\n"
            "Будут удалены все связанные алиасы.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        if data_store.remove_product(product_name):
            QMessageBox.information(self, "Готово", "Продукт удалён из справочника.")
        self._refresh_tree()

    def refresh(self):
        self._refresh_tree()


# ─────────────────────────── Публичная функция ────────────────────────────

def open_modal(parent: QWidget, app_state: dict):
    """Открывает модальный диалог отделов, блокируя родительское окно."""
    dlg = DepartmentsDialog(app_state, parent=parent)
    dlg.exec()
