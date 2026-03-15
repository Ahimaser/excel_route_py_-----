"""
departments_page.py — Управление отделами, подотделами и продуктами.

Открывается как модальный диалог (open_modal).
Отображение в стиле обозревателя решений: иконки, контекстное меню по ПКМ,
вырезание/вставка продуктов (Ctrl+X / Ctrl+V), перетаскивание в другой отдел.
"""
from __future__ import annotations

from PyQt6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTreeWidget, QTreeWidgetItem,
    QLineEdit, QComboBox, QFormLayout, QDialogButtonBox,
    QMessageBox, QInputDialog, QListWidget, QListWidgetItem, QHeaderView,
    QMenu, QApplication, QStyle, QScrollArea, QFrame,
    QToolButton,
)
from PyQt6.QtCore import Qt, QTimer, QMimeData
from PyQt6.QtGui import QFont, QAction, QShortcut, QKeySequence
from PyQt6.QtGui import QDrag, QBrush, QColor

from core import data_store
from ui.widgets import hint_icon_button, CommitLineEdit, SearchableList, ToggleSwitch

MIME_DEPT_PRODUCTS = "application/x-dept-products"


def _item_role(obj) -> tuple[str, str]:
    """Возвращает (тип, ключ/имя): 'dept'|'subdept'|'product', key или name."""
    if not isinstance(obj, dict):
        return ("", "")
    if "subdepts" in obj:
        return ("dept", obj.get("key", ""))
    if obj.get("key", "").startswith("sub_"):
        return ("subdept", obj.get("key", ""))
    return ("product", obj.get("name", ""))


class DeptsTreeWidget(QTreeWidget):
    """Дерево отделов/продуктов с перетаскиванием продуктов и контекстным меню."""

    def __init__(self, parent_dialog: "DepartmentsDialog", parent=None):
        super().__init__(parent)
        self._dialog = parent_dialog
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.setSelectionMode(QTreeWidget.SelectionMode.ExtendedSelection)
        self.setDragDropMode(QTreeWidget.DragDropMode.DragDrop)
        self.setDefaultDropAction(Qt.DropAction.MoveAction)

    def _get_selected_product_names(self) -> list[str]:
        names = []
        for item in self.selectedItems():
            obj = item.data(0, Qt.ItemDataRole.UserRole)
            kind, key = _item_role(obj)
            if kind == "product":
                names.append(key)
        return names

    def _get_target_dept_key(self, item: QTreeWidgetItem | None) -> str | None:
        if not item:
            return None
        obj = item.data(0, Qt.ItemDataRole.UserRole)
        kind, key = _item_role(obj)
        if kind in ("dept", "subdept"):
            return key
        return None

    def mimeData(self, items):
        md = QMimeData()
        names = []
        for item in items:
            obj = item.data(0, Qt.ItemDataRole.UserRole)
            kind, key = _item_role(obj)
            if kind == "product":
                names.append(key)
        if names:
            md.setData(MIME_DEPT_PRODUCTS, "\n".join(names).encode("utf-8"))
        return md

    def mimeTypes(self):
        return [MIME_DEPT_PRODUCTS]

    def dragMoveEvent(self, event):
        item = self.itemAt(event.position().toPoint())
        if self._get_target_dept_key(item):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        item = self.itemAt(event.position().toPoint())
        target_key = self._get_target_dept_key(item)
        if not target_key:
            event.ignore()
            return
        if not event.mimeData().hasFormat(MIME_DEPT_PRODUCTS):
            event.ignore()
            return
        raw = event.mimeData().data(MIME_DEPT_PRODUCTS).data().decode("utf-8")
        names = [n.strip() for n in raw.split("\n") if n.strip()]
        if names and hasattr(self._dialog, "_move_products_to"):
            self._dialog._move_products_to(names, target_key)
        event.acceptProposedAction()

    def startDrag(self, supportedActions):
        names = self._get_selected_product_names()
        if not names:
            super().startDrag(supportedActions)
            return
        md = self.mimeData(self.selectedItems())
        if not md or not md.hasFormat(MIME_DEPT_PRODUCTS):
            super().startDrag(supportedActions)
            return
        drag = QDrag(self)
        drag.setMimeData(md)
        drag.exec(Qt.DropAction.MoveAction)


class DepartmentsDialog(QDialog):
    """Модальный диалог управления отделами и продуктами."""

    def __init__(self, app_state: dict, parent=None):
        super().__init__(parent)
        self.app_state = app_state
        self.setWindowTitle("Отделы и продукты")
        self.setMinimumSize(780, 560)
        self.setModal(True)
        self._build_ui()
        self._refresh_tree()

    def _build_ui(self):
        content = QWidget()
        lay = QVBoxLayout(content)
        lay.setContentsMargins(20, 16, 20, 16)
        lay.setSpacing(12)

        title_row = QHBoxLayout()
        title_row.addWidget(QLabel("Отделы и продукты"))
        title_row.addWidget(hint_icon_button(
            self,
            "Иерархия: Отдел → Подотдел → Продукт. ПКМ — редактировать/удалить. Продукты: Ctrl+клик, Вырезать/Вставить, перетаскивание.",
            "Инструкция — Отделы и продукты\n\n"
            "1. Иерархия: Отдел → Подотделы → Продукты. Кнопки «+ Добавить» (меню) или «+ Отдел», «+ Подотдел», «+ Продукт».\n"
            "2. Бейдж [N] — количество продуктов в отделе/подотделе.\n"
            "3. ПКМ по отделу/подотделу: Переименовать, Удалить.\n"
            "4. Колонка «Печатать этикетки»: галочка включает отдел при создании этикеток.\n"
            "5. Продукты: Ctrl+клик — выбор нескольких. Ctrl+X / Ctrl+V — вырезать и вставить в другой отдел. Перетаскивание мышью.\n"
            "6. При удалении отдела продукты остаются без привязки (видны в Справочнике продуктов).",
            "Инструкция",
        ))
        title_row.addStretch()
        lay.addLayout(title_row)

        hint = QLabel(
            "Иерархия: Отдел → Подотдел → Продукт. ПКМ по элементу — редактировать или удалить. "
            "Продукты: Ctrl+клик для выбора нескольких; Ctrl+X / Ctrl+V — вырезать и вставить; перетаскивание в другой отдел."
        )
        hint.setObjectName("hintLabel")
        hint.setWordWrap(True)
        lay.addWidget(hint)

        btn_bar = QHBoxLayout()
        btn_add = QToolButton()
        btn_add.setText("+ Добавить")
        btn_add.setObjectName("btnPrimary")
        btn_add.setToolTip("Добавить отдел, подотдел или привязать продукт")
        add_menu = QMenu(self)
        act_dept = add_menu.addAction("Отдел")
        act_dept.setToolTip("Создать новый отдел")
        act_dept.triggered.connect(self._add_dept)
        act_sub = add_menu.addAction("Подотдел")
        act_sub.setToolTip("Добавить подотдел в выбранный отдел")
        act_sub.triggered.connect(self._add_subdept)
        act_prod = add_menu.addAction("Привязать продукт")
        act_prod.setToolTip("Привязать свободные продукты к отделу или подотделу")
        act_prod.triggered.connect(self._add_product)
        btn_add.setMenu(add_menu)
        btn_add.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        btn_bar.addWidget(btn_add)

        btn_add_dept = QPushButton("+ Отдел")
        btn_add_dept.setObjectName("btnSecondary")
        btn_add_dept.setToolTip("Быстро добавить отдел")
        btn_add_dept.clicked.connect(self._add_dept)
        btn_bar.addWidget(btn_add_dept)

        btn_add_subdept = QPushButton("+ Подотдел")
        btn_add_subdept.setObjectName("btnSecondary")
        btn_add_subdept.setToolTip("Добавить подотдел в выбранный отдел")
        btn_add_subdept.clicked.connect(self._add_subdept)
        btn_bar.addWidget(btn_add_subdept)

        btn_add_prod = QPushButton("+ Продукт")
        btn_add_prod.setObjectName("btnSecondary")
        btn_add_prod.setToolTip("Привязать продукт к отделу")
        btn_add_prod.clicked.connect(self._add_product)
        btn_bar.addWidget(btn_add_prod)

        btn_bar.addStretch()
        lay.addLayout(btn_bar)

        # Поиск по дереву
        search_row = QHBoxLayout()
        self.le_search = QLineEdit()
        self.le_search.setPlaceholderText("Поиск отдела или продукта...")
        self.le_search.setClearButtonEnabled(True)
        self._search_timer = QTimer(self)
        self._search_timer.setSingleShot(True)
        self._search_timer.timeout.connect(self._apply_search)
        self.le_search.textChanged.connect(lambda: self._search_timer.start(200))
        search_row.addWidget(self.le_search)
        lay.addLayout(search_row)

        self._cut_product_names: list[str] = []
        self.tree = DeptsTreeWidget(self)
        self.tree.setObjectName("deptsProductsTree")
        self.tree.setHeaderLabels(["Название", "Ед. изм. / кол-во", "Печатать этикетки"])
        hdr = self.tree.header()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.tree.setAlternatingRowColors(True)
        self.tree.setUniformRowHeights(True)
        self.tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self._show_context_menu)
        self._tree_updating = False
        lay.addWidget(self.tree)

        sc_cut = QShortcut(QKeySequence.StandardKey.Cut, self.tree)
        sc_cut.activated.connect(self._on_cut)
        sc_paste = QShortcut(QKeySequence.StandardKey.Paste, self.tree)
        sc_paste.activated.connect(self._on_paste)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_save = QPushButton("Сохранить")
        btn_save.setObjectName("btnPrimary")
        btn_save.setDefault(True)
        btn_save.setAutoDefault(True)
        btn_save.clicked.connect(self.accept)
        btn_row.addWidget(btn_save)
        lay.addLayout(btn_row)
        QShortcut(QKeySequence(Qt.Key.Key_Return), self, self.accept)

        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setWidget(content)
        main_lay = QVBoxLayout(self)
        main_lay.setContentsMargins(0, 0, 0, 0)
        main_lay.addWidget(scroll)

    # ─────────────────────────── Дерево ───────────────────────────────────

    def _build_product_index(self, products: list) -> dict:
        index: dict[str, int] = {}
        for p in products:
            k = p.get("deptKey")
            if k:
                index[k] = index.get(k, 0) + 1
        return index

    def _refresh_tree(self):
        self._tree_updating = True
        self.tree.setUpdatesEnabled(False)
        self.tree.clear()

        depts    = data_store.get_ref("departments") or []
        products = data_store.get_ref("products") or []
        prod_index = self._build_product_index(products)
        bold_font = QFont("", 13, QFont.Weight.Bold)
        style = self.tree.style()
        icon_dept = style.standardIcon(QStyle.StandardPixmap.SP_DirIcon)
        try:
            icon_sub = style.standardIcon(QStyle.StandardPixmap.SP_DirLinkIcon)
        except AttributeError:
            icon_sub = icon_dept
        icon_prod = style.standardIcon(QStyle.StandardPixmap.SP_FileIcon)

        card_bg = QBrush(QColor("#F8FAFC"))
        for dept in sorted(depts, key=lambda d: d["name"].lower()):
            dept_count = prod_index.get(dept["key"], 0)
            dept_item = QTreeWidgetItem([
                f"{dept['name']}  [{dept_count}]",
                "",
                ""
            ])
            dept_item.setData(0, Qt.ItemDataRole.UserRole, dept)
            dept_item.setIcon(0, icon_dept)
            dept_item.setFont(0, bold_font)
            dept_item.setBackground(0, card_bg)
            dept_item.setBackground(1, card_bg)
            dept_item.setBackground(2, card_bg)

            for sub in sorted(dept.get("subdepts", []), key=lambda s: s["name"].lower()):
                sub_count = prod_index.get(sub["key"], 0)
                sub_item = QTreeWidgetItem([
                    f"{sub['name']}  [{sub_count}]",
                    "",
                    ""
                ])
                sub_item.setData(0, Qt.ItemDataRole.UserRole, sub)
                sub_item.setIcon(0, icon_sub)
                sub_item.setBackground(0, card_bg)
                sub_item.setBackground(1, card_bg)
                sub_item.setBackground(2, card_bg)

                for prod in sorted(
                    (p for p in products if p.get("deptKey") == sub["key"]),
                    key=lambda p: p["name"].lower()
                ):
                    prod_item = QTreeWidgetItem([
                        prod["name"],
                        prod.get("unit", ""),
                        ""
                    ])
                    prod_item.setData(0, Qt.ItemDataRole.UserRole, prod)
                    prod_item.setIcon(0, icon_prod)
                    sub_item.addChild(prod_item)

                dept_item.addChild(sub_item)
                self._attach_labels_toggle(sub_item, bool(sub.get("labelsEnabled", True)))

            for prod in sorted(
                (p for p in products if p.get("deptKey") == dept["key"]),
                key=lambda p: p["name"].lower()
            ):
                prod_item = QTreeWidgetItem([
                    prod["name"],
                    prod.get("unit", ""),
                    ""
                ])
                prod_item.setData(0, Qt.ItemDataRole.UserRole, prod)
                prod_item.setIcon(0, icon_prod)
                dept_item.addChild(prod_item)

            self.tree.addTopLevelItem(dept_item)
            self._attach_labels_toggle(dept_item, bool(dept.get("labelsEnabled", True)))
            dept_item.setExpanded(True)

        self._tree_updating = False
        self.tree.setUpdatesEnabled(True)

    def _attach_labels_toggle(self, item: QTreeWidgetItem, enabled: bool) -> None:
        """Добавляет ToggleSwitch в колонку «Печатать этикетки» для отдела/подотдела."""
        obj = item.data(0, Qt.ItemDataRole.UserRole)
        kind, _ = _item_role(obj)
        if kind not in ("dept", "subdept"):
            return
        host = QWidget()
        row = QHBoxLayout(host)
        row.setContentsMargins(4, 2, 4, 2)
        row.setAlignment(Qt.AlignmentFlag.AlignCenter)
        tog = ToggleSwitch(host)
        tog.setChecked(enabled)
        tog.stateChanged.connect(
            lambda state, it=item: self._on_labels_toggle_changed(it, state == Qt.CheckState.Checked.value)
        )
        row.addWidget(tog)
        self.tree.setItemWidget(item, 2, host)

    def _on_labels_toggle_changed(self, item: QTreeWidgetItem, enabled: bool) -> None:
        obj = item.data(0, Qt.ItemDataRole.UserRole)
        kind, key = _item_role(obj)
        if kind not in ("dept", "subdept"):
            return
        obj["labelsEnabled"] = enabled
        depts = data_store.get_ref("departments")
        if depts:
            data_store.set_key("departments", depts)

    def _show_context_menu(self, pos):
        item = self.tree.itemAt(pos)
        if not item:
            return
        obj = item.data(0, Qt.ItemDataRole.UserRole)
        kind, key = _item_role(obj)
        menu = QMenu(self)
        if kind in ("dept", "subdept"):
            act_edit = menu.addAction("Переименовать")
            act_edit.triggered.connect(lambda: self._rename(key, kind))
            menu.addSeparator()
            act_del = menu.addAction("Удалить")
            act_del.triggered.connect(lambda: self._delete(key, kind))
            if self._cut_product_names:
                menu.addSeparator()
                act_paste = menu.addAction("Вставить")
                act_paste.triggered.connect(lambda: self._paste_into(key))
        elif kind == "product":
            act_cut = menu.addAction("Вырезать")
            act_cut.triggered.connect(lambda: self._on_cut())
            menu.addSeparator()
            act_unlink = menu.addAction("Открепить от отдела")
            act_unlink.triggered.connect(lambda: self._delete(key, kind))
            act_remove = menu.addAction("Удалить из справочника")
            act_remove.triggered.connect(lambda: self._delete_product_from_ref(key))
        menu.exec(self.tree.mapToGlobal(pos))

    def _on_cut(self):
        names = self.tree._get_selected_product_names()
        if not names:
            return
        self._cut_product_names = names
        products = data_store.get("products") or []
        for p in products:
            if p.get("name") in names:
                p["deptKey"] = None
        data_store.set_key("products", products)
        self._refresh_tree()

    def _on_paste(self):
        item = self.tree.currentItem()
        if not item or not self._cut_product_names:
            return
        target_key = self.tree._get_target_dept_key(item)
        if target_key:
            self._paste_into(target_key)

    def _paste_into(self, target_key: str):
        if not self._cut_product_names:
            return
        self._move_products_to(self._cut_product_names, target_key)
        self._cut_product_names = []
        self._refresh_tree()

    def _move_products_to(self, product_names: list[str], target_dept_key: str):
        """Переносит продукты в отдел/подотдел с ключом target_dept_key."""
        if not product_names:
            return
        products = data_store.get("products") or []
        for p in products:
            if p.get("name") in product_names:
                p["deptKey"] = target_dept_key
        data_store.set_key("products", products)
        self._refresh_tree()

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
        for d in sorted(depts, key=lambda d: d["name"].lower()):
            combo_dept.addItem(d["name"], d["key"])
        form.addRow("Отдел:", combo_dept)

        le_name = CommitLineEdit()
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
                if "очищенные" in name.lower():
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
        dlg.setMinimumHeight(500)
        vlay = QVBoxLayout(dlg)
        vlay.setSpacing(10)

        lbl_target = QLabel("Выберите отдел или подотдел для привязки:")
        vlay.addWidget(lbl_target)

        depts_list = SearchableList(dlg, multi_select=False)
        depts_list.list.setMinimumHeight(140)
        for d in sorted(depts, key=lambda d: d["name"].lower()):
            dept_item = QListWidgetItem(f"• {d['name']} (отдел)")
            dept_item.setData(Qt.ItemDataRole.UserRole, d["key"])
            depts_list.list.addItem(dept_item)
            for sub in sorted(d.get("subdepts", []), key=lambda s: s["name"].lower()):
                sub_item = QListWidgetItem(f"    └ {sub['name']} (подотдел)")
                sub_item.setData(Qt.ItemDataRole.UserRole, sub["key"])
                depts_list.list.addItem(sub_item)
        depts_list.hide_requested.connect(lambda: depts_list.setVisible(False))
        vlay.addWidget(depts_list)

        lbl_prods = QLabel("Продукты (Ctrl+клик или Shift+клик — выбрать несколько):")
        vlay.addWidget(lbl_prods)

        prods_list = SearchableList(dlg, multi_select=True)
        for p in unassigned:
            item = QListWidgetItem(f"{p['name']} ({p.get('unit', '')})")
            item.setData(Qt.ItemDataRole.UserRole, p["name"])
            prods_list.list.addItem(item)
        prods_list.hide_requested.connect(lambda: prods_list.setVisible(False))
        vlay.addWidget(prods_list)

        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        ok_btn = btns.button(QDialogButtonBox.StandardButton.Ok)
        ok_btn.setEnabled(False)

        def _update_ok():
            has_dept = depts_list.list.currentItem() is not None
            has_prods = len(prods_list.list.selectedItems()) > 0
            ok_btn.setEnabled(has_dept and has_prods)
        depts_list.list.itemSelectionChanged.connect(_update_ok)
        prods_list.list.itemSelectionChanged.connect(_update_ok)
        _update_ok()

        btns.accepted.connect(dlg.accept)
        btns.rejected.connect(dlg.reject)
        vlay.addWidget(btns)

        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        selected_names = {
            item.data(Qt.ItemDataRole.UserRole)
            for item in prods_list.list.selectedItems()
        }
        cur = depts_list.list.currentItem()
        target_key = cur.data(Qt.ItemDataRole.UserRole) if cur else None
        if not target_key or not selected_names:
            return
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

    def _apply_search(self):
        """Показывает/скрывает элементы дерева по поисковому запросу."""
        text = self.le_search.text().strip().lower()
        root = self.tree.invisibleRootItem()
        for i in range(root.childCount()):
            dept_item = root.child(i)
            dept_match = text in (dept_item.text(0) or "").lower()
            any_child_match = False
            for j in range(dept_item.childCount()):
                child = dept_item.child(j)
                child_match = text in (child.text(0) or "").lower()
                any_gc_match = False
                for k in range(child.childCount()):
                    gc = child.child(k)
                    gc_match = text in (gc.text(0) or "").lower()
                    gc.setHidden(bool(text) and not gc_match and not child_match and not dept_match)
                    if gc_match:
                        any_gc_match = True
                child_visible = dept_match or child_match or any_gc_match
                child.setHidden(bool(text) and not child_visible)
                if child_visible:
                    any_child_match = True
                    if text:
                        child.setExpanded(True)
            dept_item.setHidden(bool(text) and not dept_match and not any_child_match)
            if text and any_child_match:
                dept_item.setExpanded(True)

    def refresh(self):
        self._refresh_tree()


# ─────────────────────────── Публичная функция ────────────────────────────

def open_modal(parent: QWidget, app_state: dict):
    """Открывает модальный диалог отделов, блокируя родительское окно."""
    dlg = DepartmentsDialog(app_state, parent=parent)
    dlg.exec()
