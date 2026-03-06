"""
products_page.py — Справочник продуктов и управление алиасами.

Открывается как модальный диалог (open_modal). Отображение:

- Блок «Без отдела» показывается только после обработки файлов, если есть
  продукты без привязки (добавленные в диалоге «Новые названия» как «Новый продукт»).
  В нём можно связать вариант с каноническим (создать алиас) или привязать отдел
  в «Отделы и продукты».
- Список «Привязанные к отделам» — канонические продукты (deptKey != None).
- Таблица алиасов — связки «вариант → каноническое»; при следующем парсинге
  вариант автоматически подставляется. Кнопка «Удалить выбранные» действует по обоим спискам.
"""
from __future__ import annotations

from PyQt6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QListWidget, QListWidgetItem, QTableWidget,
    QTableWidgetItem, QHeaderView, QMessageBox,
    QGroupBox, QMenu, QLineEdit, QComboBox,
)
from PyQt6.QtCore import Qt, QMimeData
from PyQt6.QtGui import QFont, QDrag

from core import data_store
from ui.styles import STYLESHEET
from ui.widgets import hint_icon_button, make_combo_searchable

MIME_PRODUCT_NAME = "application/x-marshruty-product-name"


class LeftProductsList(QListWidget):
    """Список слева: можно перетащить элемент на правый список для связки."""
    def startDrag(self, supportedActions):
        item = self.currentItem()
        if not item:
            return
        name = item.data(Qt.ItemDataRole.UserRole)
        if not name:
            return
        mime = QMimeData()
        mime.setText(name)
        mime.setData(MIME_PRODUCT_NAME, name.encode("utf-8"))
        drag = QDrag(self)
        drag.setMimeData(mime)
        drag.exec(Qt.DropAction.CopyAction)


class RightProductsList(QListWidget):
    """Список справа: принимает перетаскивание — создаёт связку вариант → канонический."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasFormat(MIME_PRODUCT_NAME) or event.mimeData().hasText():
            event.acceptProposedAction()

    def dropEvent(self, event):
        if not event.mimeData().hasFormat(MIME_PRODUCT_NAME) and not event.mimeData().hasText():
            super().dropEvent(event)
            return
        variant = None
        if event.mimeData().hasFormat(MIME_PRODUCT_NAME):
            variant = event.mimeData().data(MIME_PRODUCT_NAME).data().decode("utf-8")
        else:
            variant = event.mimeData().text()
        item = self.itemAt(event.position().toPoint())
        if not item:
            event.ignore()
            return
        canonical = item.data(Qt.ItemDataRole.UserRole)
        if not canonical or variant == canonical:
            event.ignore()
            return
        # Связка создаётся через родительский диалог
        parent_dlg = self.window()
        if hasattr(parent_dlg, "_on_drop_link"):
            parent_dlg._on_drop_link(variant, canonical)
        event.accept()


class ProductsDialog(QDialog):
    """Модальный диалог справочника продуктов."""

    def __init__(self, app_state: dict, parent=None):
        super().__init__(parent)
        self.app_state = app_state
        self.setWindowTitle("Справочник продуктов")
        self.setMinimumSize(1200, 820)
        self.resize(1400, 920)
        self.setModal(True)
        self.setStyleSheet(STYLESHEET)
        self._build_ui()
        self._refresh()

    # ─────────────────────────── UI ───────────────────────────────────────

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 16, 20, 16)
        lay.setSpacing(14)

        title_row = QHBoxLayout()
        title_row.addWidget(QLabel("Справочник продуктов"))
        title_row.addWidget(hint_icon_button(
            self,
            "Без отдела — связка вариантов. Алиасы: вариант → каноническое. ПКМ по продукту — «Кол-во в шт».",
            "Инструкция — Справочник продуктов\n\n"
            "1. «Без отдела» (слева) — продукты без привязки к отделу. Появляется после обработки файлов.\n"
            "2. Свяжите вариант с каноническим: выберите слева и справа, нажмите «Связать» или перетащите элемент слева на правый.\n"
            "3. «Привязанные к отделам» (справа) — канонические продукты. ПКМ по продукту → «Кол-во в шт» — настройка отображения в штуках.\n"
            "4. Таблица алиасов внизу: вариант написания → каноническое. Кнопка «✕» удаляет связку.\n"
            "5. «Удалить выбранные» — удаление продуктов из справочника (по выбору в левом или правом списке).",
            "Инструкция",
        ))
        title_row.addStretch()
        lay.addLayout(title_row)

        hint = QLabel(
            "Блок «Без отдела» появляется только после обработки файлов, если вы добавили продукты как «Новый продукт». "
            "Свяжите вариант с каноническим (кнопка «Связать» или перетаскивание) — при следующих запусках это написание будет подставляться автоматически. "
            "Таблица алиасов внизу: вариант написания → каноническое название; кнопка ✕ удаляет связку."
        )
        hint.setWordWrap(True)
        hint.setObjectName("hintLabel")
        lay.addWidget(hint)

        # Поиск по названию и фильтр по отделу
        filter_row = QHBoxLayout()
        filter_row.addWidget(QLabel("Поиск по названию:"))
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Введите часть названия продукта...")
        self.search_edit.setClearButtonEnabled(True)
        self.search_edit.textChanged.connect(self._apply_filters)
        filter_row.addWidget(self.search_edit)
        filter_row.addSpacing(20)
        filter_row.addWidget(QLabel("Отдел:"))
        self.combo_dept_filter = QComboBox()
        for key, name in data_store.get_department_choices():
            self.combo_dept_filter.addItem(name, key)
        self.combo_dept_filter.currentIndexChanged.connect(self._apply_filters)
        make_combo_searchable(self.combo_dept_filter)
        filter_row.addWidget(self.combo_dept_filter)
        filter_row.addStretch()
        lay.addLayout(filter_row)

        # Левый блок (без отдела) — контейнер, видимость по данным
        lists_row = QHBoxLayout()
        lists_row.setSpacing(12)

        self.left_box = QGroupBox("Без отдела")
        left_lay = QVBoxLayout(self.left_box)
        self.list_new = LeftProductsList()
        self.list_new.setMinimumHeight(400)
        self.list_new.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        self.list_new.setAlternatingRowColors(True)
        self.list_new.setDragEnabled(True)
        self.list_new.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.list_new.customContextMenuRequested.connect(self._on_list_context_menu)
        left_lay.addWidget(self.list_new)
        lists_row.addWidget(self.left_box, 1)

        self.mid_widget = QWidget()
        mid_lay = QVBoxLayout(self.mid_widget)
        mid_lay.addStretch()
        self.btn_link = QPushButton("Связать\n→")
        self.btn_link.setObjectName("btnPrimary")
        self.btn_link.setFixedWidth(80)
        self.btn_link.clicked.connect(self._on_link)
        mid_lay.addWidget(self.btn_link)
        mid_lay.addStretch()
        lists_row.addWidget(self.mid_widget)

        right_box = QGroupBox("Привязанные к отделам (канонические)")
        right_lay = QVBoxLayout(right_box)
        self.list_canonical = RightProductsList()
        self.list_canonical.setMinimumHeight(400)
        self.list_canonical.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
        self.list_canonical.setAlternatingRowColors(True)
        self.list_canonical.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.list_canonical.customContextMenuRequested.connect(self._on_list_context_menu)
        right_lay.addWidget(self.list_canonical)
        lists_row.addWidget(right_box, 1)

        lay.addLayout(lists_row)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.btn_delete = QPushButton("Удалить выбранные из справочника")
        self.btn_delete.setObjectName("btnDanger")
        self.btn_delete.clicked.connect(self._on_delete_from_ref)
        btn_row.addWidget(self.btn_delete)
        lay.addLayout(btn_row)

        alias_box = QGroupBox("Связки (алиасы): вариант написания → каноническое название")
        alias_lay = QVBoxLayout(alias_box)

        self.alias_table = QTableWidget(0, 3)
        self.alias_table.setHorizontalHeaderLabels(
            ["Вариант написания", "→ Каноническое название", ""]
        )
        self.alias_table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.Stretch
        )
        self.alias_table.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.Stretch
        )
        self.alias_table.horizontalHeader().setSectionResizeMode(
            2, QHeaderView.ResizeMode.Fixed
        )
        self.alias_table.setColumnWidth(2, 60)
        self.alias_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.alias_table.setAlternatingRowColors(True)
        self.alias_table.setMinimumHeight(320)
        alias_lay.addWidget(self.alias_table)

        lay.addWidget(alias_box)

        btn_close = QPushButton("Закрыть")
        btn_close.setObjectName("btnSecondary")
        btn_close.clicked.connect(self._on_close_clicked)
        lay.addWidget(btn_close, alignment=Qt.AlignmentFlag.AlignRight)

    # ─────────────────────────── Данные ───────────────────────────────────

    def _refresh(self):
        """Перезагружает данные, сохраняет полные списки и применяет фильтры."""
        products = data_store.get_ref("products") or []
        aliases  = data_store.get_aliases()

        aliased_variants: set[str] = set(aliases.keys())

        self._new_prods = sorted(
            [p for p in products
             if not p.get("deptKey") and p["name"] not in aliased_variants],
            key=lambda p: p["name"].lower()
        )
        self._canonical_prods = sorted(
            [p for p in products if p.get("deptKey")],
            key=lambda p: p["name"].lower()
        )

        # Блок «Без отдела» и кнопка «Связать» — только при наличии непривязанных продуктов
        has_unassigned = len(self._new_prods) > 0
        self.left_box.setVisible(has_unassigned)
        self.mid_widget.setVisible(has_unassigned)

        self._apply_filters()

        self.alias_table.setUpdatesEnabled(False)
        self.alias_table.setRowCount(0)
        for variant, canonical in sorted(aliases.items()):
            row = self.alias_table.rowCount()
            self.alias_table.insertRow(row)
            self.alias_table.setItem(row, 0, QTableWidgetItem(variant))
            self.alias_table.setItem(row, 1, QTableWidgetItem(canonical))

            btn_del = QPushButton("✕")
            btn_del.setObjectName("btnIconDanger")
            btn_del.setFixedSize(36, 28)
            btn_del.clicked.connect(lambda _, v=variant: self._on_remove_alias(v))
            self.alias_table.setCellWidget(row, 2, btn_del)

        self.alias_table.setUpdatesEnabled(True)

    def _apply_filters(self):
        """Фильтрует списки по поиску и отделу и заполняет виджеты."""
        search = (self.search_edit.text() or "").strip().lower()
        dept_key = self.combo_dept_filter.currentData()
        if dept_key is None:
            dept_key = ""

        new_filtered = [
            p for p in (getattr(self, "_new_prods", []) or [])
            if not search or search in p["name"].lower()
        ]
        canonical_filtered = [
            p for p in (getattr(self, "_canonical_prods", []) or [])
            if (not dept_key or (p.get("deptKey") or "") == dept_key)
            and (not search or search in p["name"].lower())
        ]

        self.list_new.clear()
        for p in new_filtered:
            item = QListWidgetItem(f"{p['name']}  ({p.get('unit', '')})")
            item.setData(Qt.ItemDataRole.UserRole, p["name"])
            self.list_new.addItem(item)

        self.list_canonical.clear()
        for p in canonical_filtered:
            item = QListWidgetItem(f"{p['name']}  ({p.get('unit', '')})")
            item.setData(Qt.ItemDataRole.UserRole, p["name"])
            self.list_canonical.addItem(item)

    # ─────────────────────────── Действия ─────────────────────────────────

    def _on_drop_link(self, variant: str, canonical: str):
        """Вызывается при перетаскивании слева на правый список."""
        data_store.set_alias(variant, canonical)
        self._refresh()

    def _on_link(self):
        left_items  = self.list_new.selectedItems()
        right_items = self.list_canonical.selectedItems()

        if not left_items:
            QMessageBox.information(
                self, "Выбор",
                "Выберите одно или несколько названий в левом списке."
            )
            return
        if not right_items:
            QMessageBox.information(
                self, "Выбор",
                "Выберите каноническое название в правом списке."
            )
            return

        canonical = right_items[0].data(Qt.ItemDataRole.UserRole)
        variants  = [i.data(Qt.ItemDataRole.UserRole) for i in left_items]

        for variant in variants:
            if variant != canonical:
                data_store.set_alias(variant, canonical)

        self._refresh()

    def _has_unassigned_products(self) -> bool:
        products = data_store.get_ref("products") or []
        aliases = data_store.get_aliases()
        return any(
            p.get("name") and not p.get("deptKey") and p["name"] not in aliases
            for p in products
        )

    def _on_close_clicked(self):
        if self._has_unassigned_products():
            reply = QMessageBox.question(
                self, "Непривязанные продукты",
                "Есть продукты без отдела. Открыть окно «Отделы и продукты» для привязки?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes,
            )
            if reply == QMessageBox.StandardButton.Yes:
                self.app_state["open_departments_after_products"] = True
        self.accept()

    def _on_delete_from_ref(self):
        """Удаляет выбранные продукты из справочника (и связанные алиасы)."""
        left = [i.data(Qt.ItemDataRole.UserRole) for i in self.list_new.selectedItems()]
        right = [i.data(Qt.ItemDataRole.UserRole) for i in self.list_canonical.selectedItems()]
        names = list(dict.fromkeys(left + right))
        if not names:
            QMessageBox.information(
                self, "Выбор",
                "Выберите один или несколько продуктов в левом или правом списке."
            )
            return
        reply = QMessageBox.question(
            self, "Удалить из справочника",
            f"Удалить из справочника полностью: {len(names)} шт.?\n\n"
            "Будут удалены все связанные алиасы (связки вариантов с этим продуктом).",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        removed = 0
        for name in names:
            if data_store.remove_product(name):
                removed += 1
        self._refresh()
        if removed:
            QMessageBox.information(self, "Готово", f"Удалено из справочника: {removed}.")

    def _on_remove_alias(self, variant: str):
        data_store.remove_alias(variant)
        self._refresh()

    def _on_list_context_menu(self, pos):
        """Контекстное меню по ПКМ на элементе списка: «Кол-во в шт»."""
        list_widget = self.sender()
        if not isinstance(list_widget, QListWidget):
            return
        item = list_widget.itemAt(pos)
        if not item:
            return
        product_name = item.data(Qt.ItemDataRole.UserRole)
        if not product_name:
            return
        products = data_store.get_ref("products") or []
        prod = next((p for p in products if p.get("name") == product_name), None)
        unit = (prod.get("unit") or "").strip().lower() if prod else ""

        menu = QMenu(self)
        act_pcs = menu.addAction("Кол-во в шт")
        action = menu.exec(list_widget.mapToGlobal(pos))
        if action != act_pcs:
            return
        if unit == "шт":
            QMessageBox.information(
                self,
                "Кол-во в шт",
                "Настройка применяется только к продуктам с единицей измерения, отличной от «шт».",
            )
            return
        from ui.pages.settings_dialog import open_pcs_for_product
        open_pcs_for_product(self, product_name, on_saved=None)

    def refresh(self):
        self._refresh()


# ─────────────────────────── Публичная функция ────────────────────────────

def open_modal(parent: QWidget, app_state: dict):
    """Открывает модальный диалог справочника продуктов, блокируя родительское окно."""
    dlg = ProductsDialog(app_state, parent=parent)
    dlg.exec()
