"""
products_page.py — Справочник продуктов и управление алиасами.

Открывается как модальный диалог (open_modal).
Основное окно блокируется пока открыто это окно.

Два списка:
- Левый  («Новые / без отдела»): продукты, ещё не привязанные к отделу,
  а также варианты написания, для которых ещё нет алиаса.
- Правый («Привязанные к отделам»): продукты с deptKey != None.

Пользователь может:
1. Выбрать один или несколько продуктов из левого списка.
2. Выбрать один продукт из правого списка (каноническое название).
3. Нажать «Связать» — создаётся алиас: левый → правый.
   При следующем парсинге файлов левое написание автоматически
   заменяется на каноническое.
4. Удалить алиас кнопкой «✕» в таблице алиасов внизу.
"""
from __future__ import annotations

from PyQt6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QListWidget, QListWidgetItem, QTableWidget,
    QTableWidgetItem, QHeaderView, QMessageBox,
    QGroupBox
)
from PyQt6.QtCore import Qt, QMimeData
from PyQt6.QtGui import QFont, QDrag

from core import data_store

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
        self.setMinimumSize(820, 620)
        self.setModal(True)
        self._build_ui()
        self._refresh()

    # ─────────────────────────── UI ───────────────────────────────────────

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 16, 20, 16)
        lay.setSpacing(14)

        lbl = QLabel("Справочник продуктов")
        lbl.setObjectName("sectionTitle")
        lay.addWidget(lbl)

        hint = QLabel(
            "Левый список — новые продукты (не привязаны к отделу или не имеют алиаса). "
            "Правый список — продукты, привязанные к отделам (канонические названия). "
            "Перетащите элемент слева на нужный справа для связки или выберите и нажмите «Связать →».\n"
            "Связка объединяет разные написания одного продукта — "
            "при следующем парсинге вариант автоматически заменится на каноническое название."
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #64748b; font-size: 12px;")
        lay.addWidget(hint)

        # Два списка + кнопка связать
        lists_row = QHBoxLayout()
        lists_row.setSpacing(12)

        # Левый список
        left_box = QGroupBox("Новые / без отдела")
        left_box.setToolTip(
            "Продукты, которые ещё не привязаны к отделу\n"
            "или не имеют алиаса на каноническое название.\n"
            "Выберите один или несколько (Ctrl+клик, Shift+клик)."
        )
        left_lay = QVBoxLayout(left_box)
        self.list_new = LeftProductsList()
        self.list_new.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        self.list_new.setAlternatingRowColors(True)
        self.list_new.setDragEnabled(True)
        self.list_new.setToolTip(
            "Список новых/непривязанных продуктов.\n"
            "Можно перетащить элемент на правый список для связки.\n"
            "Ctrl+клик — выбрать несколько, Shift+клик — диапазон."
        )
        left_lay.addWidget(self.list_new)
        lists_row.addWidget(left_box, 1)

        # Кнопка связать по центру
        mid_lay = QVBoxLayout()
        mid_lay.addStretch()
        self.btn_link = QPushButton("Связать\n→")
        self.btn_link.setObjectName("btnPrimary")
        self.btn_link.setFixedWidth(80)
        self.btn_link.setToolTip(
            "Создать связку: выбранные названия слева\n"
            "→ каноническое название справа.\n"
            "При следующем парсинге вариант будет автоматически\n"
            "заменён на каноническое название."
        )
        self.btn_link.clicked.connect(self._on_link)
        mid_lay.addWidget(self.btn_link)
        mid_lay.addStretch()
        lists_row.addLayout(mid_lay)

        # Правый список
        right_box = QGroupBox("Привязанные к отделам (канонические)")
        right_box.setToolTip(
            "Продукты, привязанные к отделу/подотделу.\n"
            "Выберите одно каноническое название для создания связки."
        )
        right_lay = QVBoxLayout(right_box)
        self.list_canonical = RightProductsList()
        self.list_canonical.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
        self.list_canonical.setAlternatingRowColors(True)
        self.list_canonical.setToolTip(
            "Список канонических названий продуктов.\n"
            "Перетащите элемент из левого списка сюда для связки или выберите и нажмите «Связать»."
        )
        right_lay.addWidget(self.list_canonical)
        lists_row.addWidget(right_box, 1)

        lay.addLayout(lists_row)

        # Таблица алиасов
        alias_box = QGroupBox("Созданные связки (алиасы)")
        alias_box.setToolTip(
            "Таблица всех созданных связок вариант → каноническое название.\n"
            "Кнопка ✕ удаляет связку."
        )
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
        self.alias_table.setMaximumHeight(200)
        self.alias_table.setToolTip("Список всех созданных связок. Кнопка ✕ — удалить связку.")
        alias_lay.addWidget(self.alias_table)

        lay.addWidget(alias_box)

        btn_close = QPushButton("Закрыть")
        btn_close.setObjectName("btnSecondary")
        btn_close.setToolTip("Закрыть окно и вернуться в основное приложение")
        btn_close.clicked.connect(self._on_close_clicked)
        lay.addWidget(btn_close, alignment=Qt.AlignmentFlag.AlignRight)

    # ─────────────────────────── Данные ───────────────────────────────────

    def _refresh(self):
        """Перезагружает оба списка и таблицу алиасов."""
        products = data_store.get_ref("products") or []
        aliases  = data_store.get_aliases()

        aliased_variants: set[str] = set(aliases.keys())

        new_prods = sorted(
            [p for p in products
             if not p.get("deptKey") and p["name"] not in aliased_variants],
            key=lambda p: p["name"].lower()
        )
        canonical_prods = sorted(
            [p for p in products if p.get("deptKey")],
            key=lambda p: p["name"].lower()
        )

        self.list_new.clear()
        for p in new_prods:
            item = QListWidgetItem(f"{p['name']}  ({p.get('unit', '')})")
            item.setData(Qt.ItemDataRole.UserRole, p["name"])
            self.list_new.addItem(item)

        self.list_canonical.clear()
        for p in canonical_prods:
            item = QListWidgetItem(f"{p['name']}  ({p.get('unit', '')})")
            item.setData(Qt.ItemDataRole.UserRole, p["name"])
            self.list_canonical.addItem(item)

        self.alias_table.setUpdatesEnabled(False)
        self.alias_table.setRowCount(0)
        for variant, canonical in sorted(aliases.items()):
            row = self.alias_table.rowCount()
            self.alias_table.insertRow(row)
            self.alias_table.setItem(row, 0, QTableWidgetItem(variant))
            self.alias_table.setItem(row, 1, QTableWidgetItem(canonical))

            btn_del = QPushButton("✕")
            btn_del.setObjectName("btnIcon")
            btn_del.setStyleSheet("color: #dc2626;")
            btn_del.setFixedSize(36, 28)
            btn_del.setToolTip(f"Удалить связку «{variant}»")
            btn_del.clicked.connect(lambda _, v=variant: self._on_remove_alias(v))
            self.alias_table.setCellWidget(row, 2, btn_del)

        self.alias_table.setUpdatesEnabled(True)

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

    def _on_remove_alias(self, variant: str):
        data_store.remove_alias(variant)
        self._refresh()

    def refresh(self):
        self._refresh()


# ─────────────────────────── Публичная функция ────────────────────────────

def open_modal(parent: QWidget, app_state: dict):
    """Открывает модальный диалог справочника продуктов, блокируя родительское окно."""
    dlg = ProductsDialog(app_state, parent=parent)
    dlg.exec()
