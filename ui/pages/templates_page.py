"""
templates_page.py — Конструктор шаблонов Excel-файлов.

Функции:
- Список шаблонов с кнопками создать/удалить
- Редактор шаблона: drag-and-drop полей из панели доступных в список активных столбцов
- Двойной клик по столбцу — редактирование заголовка
- Поддержка объединённого столбца «Продукт (кол-во)»
- Открывается как модальное окно (вызывается через open_modal)
"""
from __future__ import annotations

import copy
from PyQt6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QListWidget, QListWidgetItem, QFrame, QSplitter, QInputDialog,
    QMessageBox, QLineEdit, QAbstractItemView, QMenu, QComboBox,
    QSizePolicy
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QFont, QColor

from core import data_store
from ui.styles import STYLESHEET
from ui.widgets import CommitLineEdit


# ─────────────────────────── Доступные поля ───────────────────────────────

AVAILABLE_FIELDS = [
    ("routeNumber", "№ маршрута"),
    ("address",     "Адрес"),
    ("product",     "Продукт"),
    ("unit",        "Ед. изм."),
    ("quantity",    "Количество"),
    ("pcs",         "Шт"),
    ("productQty",  "Продукт (кол-во) [объединённый]"),
]

FIELD_LABEL_MAP = {k: v for k, v in AVAILABLE_FIELDS}


# ─────────────────────────── DnD-список ───────────────────────────────────

class DragList(QListWidget):
    """QListWidget с поддержкой drag-and-drop внутри себя и из внешнего списка."""

    def __init__(self, accept_drop: bool = False, parent=None):
        super().__init__(parent)
        self.accept_drop = accept_drop
        self.setDragEnabled(True)
        self.setAcceptDrops(accept_drop)
        self.setDropIndicatorShown(accept_drop)
        if accept_drop:
            self.setDragDropMode(QAbstractItemView.DragDropMode.DragDrop)
            self.setDefaultDropAction(Qt.DropAction.MoveAction)
        else:
            self.setDragDropMode(QAbstractItemView.DragDropMode.DragOnly)
        self.setToolTip(
            "Перетащите поле в правый список, чтобы добавить его в шаблон"
            if not accept_drop else
            "Перетащите поля сюда из левого списка. Меняйте порядок перетаскиванием внутри списка"
        )

    def mimeData(self, items):
        md = super().mimeData(items)
        if items:
            item = items[0]
            field = item.data(Qt.ItemDataRole.UserRole)
            label = item.text()
            md.setText(f"{field}||{label}")
        return md

    def dropEvent(self, event):
        source = event.source()
        if source is self:
            super().dropEvent(event)
        elif source is not None:
            text = event.mimeData().text()
            if "||" in text:
                field, label = text.split("||", 1)
                self._insert_at_drop(event, field, label)
            event.acceptProposedAction()

    def _insert_at_drop(self, event, field: str, label: str):
        drop_row = self.row(self.itemAt(event.position().toPoint()))
        if drop_row < 0:
            drop_row = self.count()
        item = QListWidgetItem(label)
        item.setData(Qt.ItemDataRole.UserRole,     field)
        item.setData(Qt.ItemDataRole.UserRole + 1, None)
        item.setData(Qt.ItemDataRole.UserRole + 2, False)
        item.setData(Qt.ItemDataRole.UserRole + 3, None)
        self.insertItem(drop_row, item)


# ─────────────────────────── Диалог редактора шаблона ─────────────────────

class TemplateEditorDialog(QDialog):
    """Диалог редактирования одного шаблона."""

    def __init__(self, template: dict, unique_products: list, parent=None):
        super().__init__(parent)
        self._tmpl = copy.deepcopy(template)
        self._unique_products = unique_products
        self.setWindowTitle(f"Редактор шаблона: {template['name']}")
        self.setMinimumSize(860, 580)
        self.setModal(True)
        self._build_ui()
        self._load_columns()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(12)

        # Название
        name_row = QHBoxLayout()
        lbl_name = QLabel("Название шаблона:")
        lbl_name.setToolTip("Имя шаблона, которое отображается в списке шаблонов")
        name_row.addWidget(lbl_name)
        self.le_name = CommitLineEdit(self._tmpl["name"])
        self.le_name.setToolTip("Введите название шаблона. Изменения сохраняются по Enter или при переходе к другому полю")
        self.le_name.commit.connect(lambda: self._tmpl.update({"name": self.le_name.text().strip()}) if self.le_name.text().strip() else None)
        name_row.addWidget(self.le_name)
        root.addLayout(name_row)

        # Формат файла
        fmt_row = QHBoxLayout()
        lbl_fmt = QLabel("Формат файла:")
        lbl_fmt.setToolTip("Выберите формат генерируемого Excel-файла для этого шаблона")
        fmt_row.addWidget(lbl_fmt)
        self.combo_format = QComboBox()
        self.combo_format.setToolTip(
            "Шаблон (столбцы) — классический формат: строка маршрута + строки продуктов, \n"
            "  столбцы задаются вручную в правом списке.\n"
            "Широкий (Wide) — каждый продукт в отдельном столбце:\n"
            "  Маршрут | Адрес | ПродуктA | ПродуктB | ...\n"
            "Строчный (Rows) — строка маршрута + строки продуктов:\n"
            "  Строка маршрута: Маршрут | Адрес | — | —\n"
            "  Строка продукта: — | Название продукта | Кол-во | Шт"
        )
        self.combo_format.addItem("Шаблон (столбцы)", "")
        self.combo_format.addItem("Широкий (Wide) — продукты в столбцах", "wide")
        self.combo_format.addItem("Строчный (Rows) — продукты в строках", "rows")
        # Устанавливаем текущий формат
        current_fmt = self._tmpl.get("format", "")
        for i in range(self.combo_format.count()):
            if self.combo_format.itemData(i) == current_fmt:
                self.combo_format.setCurrentIndex(i)
                break
        self.combo_format.currentIndexChanged.connect(self._on_format_changed)
        fmt_row.addWidget(self.combo_format)
        fmt_row.addStretch()
        root.addLayout(fmt_row)

        # Подсказка о формате
        self.lbl_fmt_hint = QLabel()
        self.lbl_fmt_hint.setObjectName("infoHint")
        self.lbl_fmt_hint.setWordWrap(True)
        root.addWidget(self.lbl_fmt_hint)
        self._update_fmt_hint(current_fmt)

        hint_top = QLabel(
            "Перетащите поля из левого списка в правый. "
            "Двойной клик по столбцу — изменить заголовок. "
            "Правый клик — дополнительные действия."
        )
        hint_top.setObjectName("hintLabel")
        hint_top.setWordWrap(True)
        root.addWidget(hint_top)

        # Основная область
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # Левая панель — доступные поля
        left = QWidget()
        left_lay = QVBoxLayout(left)
        left_lay.setContentsMargins(0, 0, 0, 0)
        lbl_avail = QLabel("Доступные поля:")
        lbl_avail.setToolTip("Перетащите нужное поле в правый список")
        left_lay.addWidget(lbl_avail)
        self.list_available = DragList(accept_drop=False)
        for field, label in AVAILABLE_FIELDS:
            item = QListWidgetItem(label)
            item.setData(Qt.ItemDataRole.UserRole, field)
            item.setToolTip(self._field_hint(field))
            self.list_available.addItem(item)
        left_lay.addWidget(self.list_available)
        splitter.addWidget(left)

        # Правая панель — активные столбцы
        right = QWidget()
        right_lay = QVBoxLayout(right)
        right_lay.setContentsMargins(0, 0, 0, 0)
        lbl_active = QLabel("Активные столбцы (двойной клик — изменить заголовок):")
        lbl_active.setToolTip("Столбцы, которые будут в итоговом Excel-файле. Порядок — перетаскивание")
        right_lay.addWidget(lbl_active)
        self.list_active = DragList(accept_drop=True)
        self.list_active.itemDoubleClicked.connect(self._on_rename_column)
        self.list_active.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.list_active.customContextMenuRequested.connect(self._on_active_context_menu)
        right_lay.addWidget(self.list_active)

        btn_remove = QPushButton("Удалить выбранный столбец")
        btn_remove.setObjectName("btnDanger")
        btn_remove.setToolTip("Удалить выделенный столбец из шаблона")
        btn_remove.clicked.connect(self._remove_selected)
        right_lay.addWidget(btn_remove)
        splitter.addWidget(right)

        splitter.setSizes([300, 500])
        root.addWidget(splitter)

        # Кнопки
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_cancel = QPushButton("Отмена")
        btn_cancel.setObjectName("btnSecondary")
        btn_cancel.setToolTip("Закрыть без сохранения изменений")
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_cancel)
        btn_save = QPushButton("Сохранить")
        btn_save.setObjectName("btnPrimary")
        btn_save.setToolTip("Сохранить изменения шаблона")
        btn_save.clicked.connect(self._on_save)
        btn_row.addWidget(btn_save)
        root.addLayout(btn_row)

    @staticmethod
    def _field_hint(field: str) -> str:
        hints = {
            "routeNumber": "Номер маршрута из исходного файла",
            "address":     "Адрес маршрута",
            "product":     "Название продукта",
            "unit":        "Единица измерения (кг, шт и т.д.)",
            "quantity":    "Количество продукта",
            "pcs":         "Количество в штуках (рассчитывается по настройкам Шт)",
            "productQty":  "Объединённый столбец: заголовок — название продукта, строки — количество. "
                           "Используется когда в подотделе только один продукт.",
        }
        return hints.get(field, "")

    def _on_format_changed(self, _index: int) -> None:
        """Обновляет подсказку при смене формата и показывает/скрывает список столбцов."""
        fmt = self.combo_format.currentData() or ""
        self._tmpl["format"] = fmt
        self._update_fmt_hint(fmt)
        # Для wide/rows список столбцов не используется
        is_legacy = fmt == ""
        # Показываем/скрываем splitter со списками столбцов
        # Ищем splitter в layout
        for i in range(self.layout().count()):
            item = self.layout().itemAt(i)
            if item and item.widget():
                from PyQt6.QtWidgets import QSplitter
                if isinstance(item.widget(), QSplitter):
                    item.widget().setVisible(is_legacy)
                    break

    def _update_fmt_hint(self, fmt: str) -> None:
        """Обновляет текст подсказки о выбранном формате."""
        hints = {
            "": (
                "Классический формат: строки маршрутов и продуктов, столбцы задаются вручную "
                "в правом списке. Поддерживает все поля включая объединённый столбец."
            ),
            "wide": (
                "Широкий формат: каждый продукт в отдельном столбце. "
                "Маршрут | Адрес | Продукт1 | Продукт2 | ... "
                "Значение ячейки: \"5 кг / 3 шт\" (шт — опционально). "
                "Список столбцов не используется."
            ),
            "rows": (
                "Строчный формат: строка маршрута (Маршрут | Адрес | — | —) "
                "+ строки продуктов (— | Название продукта | Кол-во | Шт). "
                "Список столбцов не используется."
            ),
        }
        self.lbl_fmt_hint.setText(hints.get(fmt, ""))

    def _load_columns(self):
        self.list_active.clear()
        for col in self._tmpl.get("columns", []):
            field = col["field"]
            custom_label = col.get("label")
            merged = col.get("merged", False)
            product_name = col.get("productName")

            if custom_label:
                display = custom_label
            elif merged and product_name:
                display = product_name
            else:
                display = FIELD_LABEL_MAP.get(field, field)

            item = QListWidgetItem(display)
            item.setData(Qt.ItemDataRole.UserRole,     field)
            item.setData(Qt.ItemDataRole.UserRole + 1, custom_label)
            item.setData(Qt.ItemDataRole.UserRole + 2, merged)
            item.setData(Qt.ItemDataRole.UserRole + 3, product_name)
            item.setToolTip(self._field_hint(field))
            if merged:
                item.setForeground(QColor("#2563eb"))
                f = item.font(); f.setItalic(True); item.setFont(f)
            self.list_active.addItem(item)

    def _on_rename_column(self, item: QListWidgetItem):
        current = item.data(Qt.ItemDataRole.UserRole + 1) or item.text()
        new_label, ok = QInputDialog.getText(
            self, "Изменить заголовок",
            "Введите новый заголовок столбца\n(не влияет на содержимое — только название в Excel):",
            text=current
        )
        if ok and new_label.strip():
            item.setData(Qt.ItemDataRole.UserRole + 1, new_label.strip())
            item.setText(new_label.strip())

    def _on_active_context_menu(self, pos):
        item = self.list_active.itemAt(pos)
        if not item:
            return
        field = item.data(Qt.ItemDataRole.UserRole)
        menu = QMenu(self)
        menu.setToolTipsVisible(True)

        if field == "productQty":
            act_prod = menu.addAction("Выбрать продукт для столбца...")
            act_prod.setToolTip("Выбрать продукт, количество которого будет в этом столбце")
            menu.addSeparator()

        act_reset = menu.addAction("Сбросить заголовок к умолчанию")
        act_reset.setToolTip("Вернуть стандартное название заголовка")
        act_del = menu.addAction("Удалить столбец")
        act_del.setToolTip("Удалить этот столбец из шаблона")

        action = menu.exec(self.list_active.viewport().mapToGlobal(pos))
        if action is None:
            return

        if field == "productQty" and action == act_prod:
            self._pick_product_for_column(item)
        elif action == act_reset:
            item.setData(Qt.ItemDataRole.UserRole + 1, None)
            merged = item.data(Qt.ItemDataRole.UserRole + 2)
            pname  = item.data(Qt.ItemDataRole.UserRole + 3)
            item.setText(pname if (merged and pname) else FIELD_LABEL_MAP.get(field, field))
        elif action == act_del:
            self.list_active.takeItem(self.list_active.row(item))

    def _pick_product_for_column(self, item: QListWidgetItem):
        if not self._unique_products:
            QMessageBox.information(
                self, "Нет продуктов",
                "Сначала загрузите файлы с маршрутами на главной странице.\n"
                "Продукт будет определён автоматически при наличии данных."
            )
            return
        dlg = QInputDialog(self)
        dlg.setWindowTitle("Выбор продукта")
        dlg.setLabelText(
            "Выберите продукт для объединённого столбца.\n"
            "Заголовок столбца = название продукта, строки = количество:"
        )
        dlg.setComboBoxItems(sorted(self._unique_products))
        dlg.setComboBoxEditable(False)
        if dlg.exec() == QInputDialog.DialogCode.Accepted:
            pname = dlg.textValue()
            item.setData(Qt.ItemDataRole.UserRole + 2, True)
            item.setData(Qt.ItemDataRole.UserRole + 3, pname)
            custom = item.data(Qt.ItemDataRole.UserRole + 1)
            item.setText(custom if custom else pname)
            item.setForeground(QColor("#2563eb"))
            f = item.font(); f.setItalic(True); item.setFont(f)

    def _remove_selected(self):
        for item in self.list_active.selectedItems():
            self.list_active.takeItem(self.list_active.row(item))

    def _on_save(self):
        name = self.le_name.text().strip()
        if not name:
            QMessageBox.warning(self, "Ошибка", "Введите название шаблона.")
            return

        fmt = self.combo_format.currentData() or ""
        columns = []

        # Для wide/rows список столбцов не требуется
        if fmt == "":
            if self.list_active.count() == 0:
                QMessageBox.warning(self, "Ошибка", "Добавьте хотя бы один столбец.")
                return
            for i in range(self.list_active.count()):
                item = self.list_active.item(i)
                field        = item.data(Qt.ItemDataRole.UserRole)
                custom_label = item.data(Qt.ItemDataRole.UserRole + 1)
                merged       = item.data(Qt.ItemDataRole.UserRole + 2) or False
                product_name = item.data(Qt.ItemDataRole.UserRole + 3)
                col = {"field": field, "label": custom_label, "merged": merged}
                if merged and product_name:
                    col["productName"] = product_name
                columns.append(col)

        data_store.save_template(
            self._tmpl["id"], name, columns,
            self._tmpl.get("deptKey"), fmt
        )
        self.accept()


# ─────────────────────────── Главный диалог шаблонов ──────────────────────

class TemplatesDialog(QDialog):
    """Модальный диалог управления шаблонами."""

    def __init__(self, app_state: dict, parent=None):
        super().__init__(parent)
        self.app_state = app_state
        self.setWindowTitle("Шаблоны")
        self.setMinimumSize(640, 440)
        self.setModal(True)
        self.setStyleSheet(STYLESHEET)
        self._build_ui()
        self._refresh_list()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(20, 16, 20, 16)
        root.setSpacing(12)

        lbl_title = QLabel("Управление шаблонами Excel-файлов")
        lbl_title.setObjectName("sectionTitle")
        root.addWidget(lbl_title)

        hint = QLabel(
            "Шаблон определяет набор и порядок столбцов в файлах по отделам. "
            "Двойной клик по шаблону — открыть редактор."
        )
        hint.setObjectName("hintLabel")
        hint.setWordWrap(True)
        root.addWidget(hint)

        self.list_templates = QListWidget()
        self.list_templates.setAlternatingRowColors(True)
        self.list_templates.setToolTip(
            "Список шаблонов. Двойной клик — редактировать. "
            "Выберите шаблон и нажмите кнопку для действия."
        )
        self.list_templates.itemDoubleClicked.connect(self._on_edit)
        root.addWidget(self.list_templates)

        btn_row = QHBoxLayout()
        btn_new = QPushButton("Создать шаблон")
        btn_new.setObjectName("btnPrimary")
        btn_new.setToolTip("Создать новый пустой шаблон")
        btn_new.clicked.connect(self._on_create)
        btn_row.addWidget(btn_new)

        btn_edit = QPushButton("Редактировать")
        btn_edit.setObjectName("btnSecondary")
        btn_edit.setToolTip("Открыть редактор выбранного шаблона")
        btn_edit.clicked.connect(self._on_edit)
        btn_row.addWidget(btn_edit)

        btn_del = QPushButton("Удалить")
        btn_del.setObjectName("btnDanger")
        btn_del.setToolTip("Удалить выбранный шаблон (нельзя удалить последний)")
        btn_del.clicked.connect(self._on_delete)
        btn_row.addWidget(btn_del)

        btn_row.addStretch()
        btn_close = QPushButton("Закрыть")
        btn_close.setObjectName("btnSecondary")
        btn_close.setToolTip("Закрыть окно шаблонов")
        btn_close.clicked.connect(self.accept)
        btn_row.addWidget(btn_close)
        root.addLayout(btn_row)

    def _refresh_list(self):
        self.list_templates.clear()
        templates = data_store.get_ref("templates") or []
        for tmpl in templates:
            cols = tmpl.get("columns", [])
            col_names = []
            for c in cols:
                lbl = (c.get("label")
                       or (c.get("productName") if c.get("merged") else None)
                       or data_store.FIELD_LABELS.get(c["field"], c["field"]))
                col_names.append(lbl)
            summary = " | ".join(col_names[:6])
            if len(col_names) > 6:
                summary += f" (+{len(col_names)-6})"
            item = QListWidgetItem(f"{tmpl['name']}\n  Столбцы: {summary}")
            item.setData(Qt.ItemDataRole.UserRole, tmpl["id"])
            self.list_templates.addItem(item)

    def _get_selected_id(self) -> str | None:
        items = self.list_templates.selectedItems()
        return items[0].data(Qt.ItemDataRole.UserRole) if items else None

    def _on_create(self):
        name, ok = QInputDialog.getText(
            self, "Новый шаблон",
            "Введите название нового шаблона:"
        )
        if not ok or not name.strip():
            return
        tmpl = data_store.create_template(name.strip())
        self._refresh_list()
        self._edit_template(tmpl["id"])

    def _on_edit(self):
        tid = self._get_selected_id()
        if not tid:
            QMessageBox.information(self, "Выберите шаблон",
                                    "Выберите шаблон из списка для редактирования.")
            return
        self._edit_template(tid)

    def _edit_template(self, template_id: str):
        templates = data_store.get_ref("templates") or []
        tmpl = next((t for t in templates if t["id"] == template_id), None)
        if not tmpl:
            return
        unique_prods = [p["name"] for p in (self.app_state.get("uniqueProducts") or [])]
        dlg = TemplateEditorDialog(tmpl, unique_prods, parent=self)
        dlg.exec()
        self._refresh_list()

    def _on_delete(self):
        tid = self._get_selected_id()
        if not tid:
            QMessageBox.information(self, "Выберите шаблон",
                                    "Выберите шаблон из списка для удаления.")
            return
        templates = data_store.get_ref("templates") or []
        if len(templates) <= 1:
            QMessageBox.warning(self, "Нельзя удалить",
                                "Должен остаться хотя бы один шаблон.")
            return
        tmpl = next((t for t in templates if t["id"] == tid), None)
        name = tmpl["name"] if tmpl else tid
        reply = QMessageBox.question(
            self, "Удалить шаблон",
            f"Удалить шаблон «{name}»?\nЭто действие нельзя отменить.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            data_store.delete_template(tid)
            self._refresh_list()


# ─────────────────────────── Публичная функция ────────────────────────────

def open_modal(parent: QWidget, app_state: dict):
    """Открывает модальный диалог шаблонов, блокируя родительское окно."""
    dlg = TemplatesDialog(app_state, parent=parent)
    dlg.exec()
