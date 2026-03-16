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
    QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox,
    QGroupBox, QLineEdit, QComboBox, QScrollArea, QFrame,
    QTabWidget, QMenu, QWidgetAction, QInputDialog, QStyle,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QShortcut, QKeySequence

from core import data_store
from ui.widgets import hint_icon_button


class ProductsDialog(QDialog):
    """Модальный диалог справочника продуктов."""

    def __init__(self, app_state: dict, parent=None):
        super().__init__(parent)
        self.app_state = app_state
        self.setWindowTitle("Справочник продуктов")
        self.setMinimumSize(1200, 820)
        self.resize(1400, 920)
        self.setModal(True)
        self._build_ui()
        self._refresh()

    # ─────────────────────────── UI ───────────────────────────────────────

    def _build_ui(self):
        content = QWidget()
        lay = QVBoxLayout(content)
        lay.setContentsMargins(20, 16, 20, 16)
        lay.setSpacing(14)

        title_row = QHBoxLayout()
        title_row.addWidget(QLabel("Справочник продуктов"))
        title_row.addWidget(hint_icon_button(
            self,
            "Вкладки: Все, Без отдела. Алиасы связывают вариант написания с каноническим названием.",
            "Инструкция — Справочник продуктов\n\n"
            "1. Вкладки: «Все», «Без отдела».\n"
            "2. Для продуктов без отдела: «В отдел» — привязка к отделу, «Связать с» — создать алиас (вариант → каноническое).\n"
            "3. Таблица алиасов внизу: вариант написания (каноническое). Кнопка удаления — удаляет связку.\n"
            "4. Значок корзины в строке — удаление продукта из справочника.\n"
            "5. Настройки количества в штуках — меню «Настройки» → «Настройки Количества».",
            "Инструкция",
        ))
        title_row.addStretch()
        lay.addLayout(title_row)

        hint = QLabel(
            "Вкладки «Все», «Без отдела». Для продуктов без отдела — «В отдел» (привязка) и «Связать с» (алиас). "
            "Таблица алиасов внизу: вариант написания → каноническое название."
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
        filter_row.addWidget(self.combo_dept_filter)
        filter_row.addStretch()
        btn_add = QPushButton("+ Добавить продукт")
        btn_add.setObjectName("btnPrimary")
        btn_add.clicked.connect(self._on_add_product)
        filter_row.addWidget(btn_add)
        lay.addLayout(filter_row)

        # Единая таблица с вкладками
        self.tab_widget = QTabWidget()
        for _ in range(2):
            tab = QWidget()
            tab.setLayout(QVBoxLayout())
            tab.layout().setContentsMargins(0, 0, 0, 0)
            self.tab_widget.addTab(tab, "")
        self.tab_widget.setTabText(0, "Все")
        self.tab_widget.setTabText(1, "Без отдела")
        self.tab_widget.currentChanged.connect(self._on_tab_changed)

        self.products_table = QTableWidget(0, 4)
        self.products_table.setHorizontalHeaderLabels(["Название", "Отдел", "Ед. изм.", "Действия"])
        self.products_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.products_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.products_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.products_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)
        self.products_table.setColumnWidth(3, 300)
        self.products_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.products_table.setAlternatingRowColors(True)
        self.products_table.setMinimumHeight(400)
        self.products_table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        self.products_table.verticalHeader().setDefaultSectionSize(48)

        self.tab_widget.widget(0).layout().addWidget(self.products_table)
        lay.addWidget(self.tab_widget)

        alias_box = QGroupBox("Связки (алиасы): вариант написания → каноническое название")
        alias_lay = QVBoxLayout(alias_box)

        self.alias_table = QTableWidget(0, 2)
        self.alias_table.setHorizontalHeaderLabels(
            ["Вариант (каноническое)", ""]
        )
        self.alias_table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.Stretch
        )
        self.alias_table.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.Fixed
        )
        self.alias_table.setColumnWidth(1, 60)
        self.alias_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.alias_table.setAlternatingRowColors(True)
        self.alias_table.setMinimumHeight(320)
        alias_lay.addWidget(self.alias_table)

        lay.addWidget(alias_box)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_save = QPushButton("Сохранить")
        btn_save.setObjectName("btnPrimary")
        btn_save.setDefault(True)
        btn_save.setAutoDefault(True)
        btn_save.clicked.connect(self._on_close_clicked)
        btn_row.addWidget(btn_save)
        lay.addLayout(btn_row)
        QShortcut(QKeySequence(Qt.Key.Key_Return), self, self._on_close_clicked)

        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setWidget(content)
        main_lay = QVBoxLayout(self)
        main_lay.setContentsMargins(0, 0, 0, 0)
        main_lay.addWidget(scroll)

    # ─────────────────────────── Данные ───────────────────────────────────

    def _on_tab_changed(self, index: int):
        """При смене вкладки — переключаем таблицу и обновляем данные."""
        tab = self.tab_widget.widget(index)
        if self.products_table.parent() != tab:
            self.products_table.setParent(tab)
            tab.layout().addWidget(self.products_table)
        self._apply_filters()

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

        self._apply_filters()

        self.alias_table.setUpdatesEnabled(False)
        self.alias_table.setRowCount(0)
        for variant, canonical in sorted(aliases.items()):
            row = self.alias_table.rowCount()
            self.alias_table.insertRow(row)
            self.alias_table.setItem(row, 0, QTableWidgetItem(f"{variant} ({canonical})"))

            icon_del = self.style().standardIcon(QStyle.StandardPixmap.SP_TrashIcon)
            btn_del = QPushButton(icon_del, "")
            btn_del.setObjectName("btnIconDanger")
            btn_del.setFixedSize(36, 28)
            btn_del.setToolTip("Удалить связку")
            btn_del.clicked.connect(lambda _, v=variant: self._on_remove_alias(v))
            self.alias_table.setCellWidget(row, 1, btn_del)

        self.alias_table.setUpdatesEnabled(True)

    def _apply_filters(self):
        """Фильтрует по вкладке, поиску и отделу, заполняет таблицу."""
        search = (self.search_edit.text() or "").strip().lower()
        dept_key = self.combo_dept_filter.currentData()
        if dept_key is None:
            dept_key = ""
        tab_idx = self.tab_widget.currentIndex()

        new_filtered = [
            p for p in (getattr(self, "_new_prods", []) or [])
            if not search or search in p["name"].lower()
        ]
        canonical_filtered = [
            p for p in (getattr(self, "_canonical_prods", []) or [])
            if (not dept_key or (p.get("deptKey") or "") == dept_key)
            and (not search or search in p["name"].lower())
        ]

        if tab_idx == 0:
            rows_data = new_filtered + canonical_filtered
        else:
            rows_data = new_filtered

        rows_data.sort(key=lambda p: p["name"].lower())

        self._populate_products_table(rows_data)

    def _populate_products_table(self, rows_data: list):
        """Заполняет таблицу продуктов. rows_data: список {name, unit, deptKey?}."""
        canonical_names = {p["name"] for p in (getattr(self, "_canonical_prods", []) or [])}
        prod_map = data_store.get_products_map()

        self.products_table.setUpdatesEnabled(False)
        self.products_table.setRowCount(0)

        for p in rows_data:
            row = self.products_table.rowCount()
            self.products_table.insertRow(row)
            name = p.get("name", "")
            unit = p.get("unit", "")
            dept_key = p.get("deptKey") or ""
            is_unassigned = not dept_key

            # Название: при showPcs — «название (кол-во в 1 шт + ед. изм.)», иначе «название (ед. изм.)»
            sp = prod_map.get(name, {})
            if sp.get("showPcs") and unit and unit.lower() != "шт":
                name_txt = data_store.format_product_display_name(name, prod_map)
            else:
                name_txt = f"{name} ({unit})" if unit else name
            name_item = QTableWidgetItem(name_txt)
            name_item.setData(Qt.ItemDataRole.UserRole, name)
            self.products_table.setItem(row, 0, name_item)

            dept_txt = data_store.get_department_display_name(dept_key) if dept_key else "—"
            self.products_table.setItem(row, 1, QTableWidgetItem(dept_txt))
            self.products_table.setItem(row, 2, QTableWidgetItem(unit or ""))

            # Действия
            actions_widget = QWidget()
            actions_lay = QHBoxLayout(actions_widget)
            actions_lay.setContentsMargins(2, 4, 2, 4)
            actions_lay.setSpacing(4)

            if is_unassigned:
                btn_dept = QPushButton("В отдел")
                btn_dept.setObjectName("btnPrimary")
                btn_dept.setFixedHeight(28)
                btn_dept.clicked.connect(lambda _, n=name, b=btn_dept: self._on_assign_dept(n, b))
                actions_lay.addWidget(btn_dept)

                combo_link = QComboBox()
                combo_link.addItem("Связать с ▼", None)
                for c in canonical_names:
                    if c != name:
                        combo_link.addItem(c, c)
                combo_link.currentIndexChanged.connect(
                    lambda idx, n=name, cb=combo_link: self._on_link_combo_changed(n, cb)
                )
                combo_link.setFixedHeight(28)
                combo_link.setMinimumWidth(140)
                actions_lay.addWidget(combo_link)

            if unit and unit.lower() != "шт":
                btn_variant = QPushButton("Вариант")
                btn_variant.setObjectName("btnSecondary")
                btn_variant.setFixedHeight(28)
                btn_variant.setToolTip("Создать вариант с другим количеством в 1 шт/коробке")
                btn_variant.clicked.connect(lambda _, n=name, u=unit: self._on_add_variant(n, u, dept_key))
                actions_lay.addWidget(btn_variant)

            icon_del = self.style().standardIcon(QStyle.StandardPixmap.SP_TrashIcon)
            btn_del = QPushButton(icon_del, "")
            btn_del.setObjectName("btnIconDanger")
            btn_del.setFixedSize(36, 32)
            btn_del.setToolTip("Удалить продукт из справочника")
            btn_del.clicked.connect(lambda _, n=name: self._on_delete_product(n))
            actions_lay.addWidget(btn_del)

            actions_lay.addStretch()
            self.products_table.setCellWidget(row, 3, actions_widget)

        self.products_table.setUpdatesEnabled(True)

    # ─────────────────────────── Действия ─────────────────────────────────

    def _on_add_product(self):
        """Добавить продукт вручную в справочник (карточная форма)."""
        from ui.pages.add_product_dialog import open_add_product_dialog
        if open_add_product_dialog(self):
            self._refresh()
            QMessageBox.information(self, "Готово", "Продукт добавлен в справочник.")

    def _on_add_variant(self, base_name: str, unit: str, dept_key: str):
        """Создать вариант продукта с другим количеством в 1 шт/коробке (напр. масло сливочное (0,18 кг))."""
        base_prod = next((p for p in (data_store.get_ref("products") or []) if p.get("name") == base_name), None)
        base_pcu = float(base_prod.get("pcsPerUnit", 1) or 1) if base_prod else 1.0
        pcu, ok = QInputDialog.getDouble(
            self, "Добавить вариант продукта",
            f"Количество в 1 шт/коробке для «{base_name}»:\n(будет создан продукт «{base_name} (X {unit})»)",
            base_pcu, 0.001, 999.999, 3,
        )
        if not ok or pcu <= 0:
            return
        pcu_str = str(pcu).replace(".", ",")
        variant_name = f"{base_name} ({pcu_str} {unit})"
        if any(p.get("name") == variant_name for p in (data_store.get_ref("products") or [])):
            QMessageBox.warning(
                self, "Ошибка",
                f"Продукт «{variant_name}» уже есть в справочнике.",
            )
            return
        if data_store.add_product(variant_name, unit, dept_key or None):
            data_store.update_product(variant_name, showPcs=True, pcsPerUnit=pcu)
            self._refresh()
            QMessageBox.information(
                self, "Готово",
                f"Вариант «{variant_name}» добавлен с настройкой Шт: {pcu_str} {unit} в 1 шт.",
            )
        else:
            QMessageBox.warning(self, "Ошибка", f"Не удалось добавить «{variant_name}».")

    def _on_assign_dept(self, product_name: str, button: QPushButton):
        """Кнопка «В отдел» — меню выбора отдела."""
        menu = QMenu(self)
        for key, display_name in data_store.get_department_choices():
            if not key:
                continue
            action = menu.addAction(display_name)
            action.triggered.connect(
                lambda checked=False, k=key: self._do_assign_dept(product_name, k)
            )
        menu.exec(button.mapToGlobal(button.rect().bottomLeft()))

    def _do_assign_dept(self, product_name: str, dept_key: str):
        data_store.update_product(product_name, deptKey=dept_key)
        self._refresh()

    def _on_link_combo_changed(self, variant: str, combo: QComboBox):
        canonical = combo.currentData()
        if canonical and variant != canonical:
            data_store.set_alias(variant, canonical)
            self._refresh()

    def _on_delete_product(self, name: str):
        reply = QMessageBox.question(
            self, "Удалить из справочника",
            f"Удалить «{name}» из справочника?\n\nБудут удалены все связанные алиасы.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes and data_store.remove_product(name):
            self._refresh()
            QMessageBox.information(self, "Готово", "Удалено из справочника.")

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
