"""
templates_page.py — Конструктор шаблонов Excel-файлов.

- Окно создания/редактирования: таблица с настраиваемым размером.
- Строка 1 по умолчанию: «Маршруты» + отдел + дата + основной/увеличение.
- Перетаскивание полей в ячейки, объединение ячеек (ПКМ).
"""
from __future__ import annotations

import copy
from datetime import datetime
from PyQt6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QListWidget, QListWidgetItem, QSplitter, QTableWidget,
    QTableWidgetItem, QMessageBox, QLineEdit, QAbstractItemView, QMenu,
    QComboBox, QHeaderView, QInputDialog, QSpinBox,
    QDateEdit, QScrollArea, QGridLayout, QFrame,
    QTextBrowser, QGroupBox, QCheckBox,
)
from PyQt6.QtCore import Qt, pyqtSignal, QMimeData, QDate, QPoint
from PyQt6.QtGui import QColor, QDrag, QPixmap

from core import data_store
from ui.widgets import CommitLineEdit, hint_icon_button, ToggleSwitch


# ─────────────────────────── Доступные поля ───────────────────────────────

AVAILABLE_FIELDS = [
    ("routeNumber", "№ маршрута"),
    ("address", "Адрес"),
    ("product", "Продукт"),
    ("unit", "Ед. изм."),
    ("quantity", "Количество"),
    ("pcs", "Шт"),
    ("productQty", "Продукт (кол-во)"),
    ("productsWide", "Продукт (колонка на каждый)"),
    ("nomenclature", "Номенклатура"),
]

# Пояснения к полям (показываются как подсказка при наведении на элемент)
FIELD_DESCRIPTIONS: dict[str, str] = {
    "routeNumber": "Номер маршрута. В шапке — заголовок колонки, в данных — номер маршрута по каждому адресу/продукту.",
    "address": "Адрес доставки. В шапке — заголовок, в данных — адрес точки маршрута.",
    "product": "Название продукта. В шапке — заголовок, в данных — наименование продукта.",
    "unit": "Единица измерения (кг, л, шт и т.д.). В шапке — заголовок, в данных — ед. изм. продукта.",
    "quantity": "Количество в выбранных единицах (с учётом коэффициента замены). В шапке — заголовок, в данных — число.",
    "pcs": "Количество в штуках (для маршрутов ШК/СД — округлённое). В шапке — «Шт», в данных — число или «—».",
    "productQty": "Одна колонка: в шапке — заголовок, в данных — название продукта и его количество в одной ячейке.",
    "productsWide": "Отдельная колонка на каждый продукт отдела/подотдела: в шапке — название продукта, в данных — только количество по этому продукту.",
    "nomenclature": "В шапке — «Номенклатура». В данных: первая строка по маршруту — адрес; следующие строки — продукты этого маршрута (название и количество). Номер маршрута выводится только в строке с адресом.",
}

FIELD_LABEL_MAP = {k: v for k, v in AVAILABLE_FIELDS}
MIME_FIELD = "application/x-template-field"
MIME_FIELD_FROM_CELL = "application/x-template-field-from-cell"  # перетаскивание из ячейки (move)


# Подставляется при создании файла по типу загруженного маршрута (основной/довоз)
TITLE_TYPE_PLACEHOLDER = "основной/увеличение"


def _build_title_string(include_dept: bool, dept_name: str, date_str: str, type_str: str | None = None) -> str:
    """Собирает строку заголовка. type_str=None — подставляется placeholder (определится при создании файла)."""
    parts = ["Маршруты"]
    if include_dept and dept_name:
        parts.append(dept_name)
    if date_str:
        parts.append(date_str)
    parts.append("основной" if type_str == "main" else ("увеличение" if type_str == "increase" else TITLE_TYPE_PLACEHOLDER))
    return " ".join(parts)


def _parse_date_to_qdate(s: str) -> QDate:
    """Парсит DD.MM.YYYY в QDate."""
    try:
        dt = datetime.strptime(s.strip(), "%d.%m.%Y")
        return QDate(dt.year, dt.month, dt.day)
    except (ValueError, TypeError):
        from datetime import date, timedelta
        t = date.today() + timedelta(days=1)
        return QDate(t.year, t.month, t.day)


# ─────────────────────────── Поля в несколько столбцов (компактно) ─────────

class DraggableFieldButton(QPushButton):
    """Кнопка поля для перетаскивания в таблицу."""

    def __init__(self, field: str, label: str, parent=None):
        super().__init__(label, parent)
        self._field = field
        self._label = label
        self.setFlat(True)
        self.setFixedHeight(24)
        self.setStyleSheet("font-size: 11px; text-align: left; padding: 2px 6px;")
        if field in FIELD_DESCRIPTIONS:
            self.setToolTip(FIELD_DESCRIPTIONS[field])
        self._drag_start = None

    def mousePressEvent(self, event):
        self._drag_start = event.position().toPoint() if hasattr(event, "position") else event.pos()
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._drag_start is None:
            super().mouseMoveEvent(event)
            return
        pos = event.position().toPoint() if hasattr(event, "position") else event.pos()
        if (pos - self._drag_start).manhattanLength() < 8:
            super().mouseMoveEvent(event)
            return
        md = QMimeData()
        md.setData(MIME_FIELD, f"{self._field}||{self._label}".encode("utf-8"))
        md.setText(f"{self._field}||{self._label}")
        drag = QDrag(self)
        drag.setMimeData(md)
        # Превью перетаскиваемого элемента
        pixmap = QPixmap(self.size())
        pixmap.fill(Qt.GlobalColor.transparent)
        self.render(pixmap)
        drag.setPixmap(pixmap)
        drag.setHotSpot(QPoint(self.width() // 2, self.height() // 2))
        drag.exec(Qt.DropAction.CopyAction)
        self._drag_start = None

    def mouseReleaseEvent(self, event):
        self._drag_start = None
        super().mouseReleaseEvent(event)


class FieldsGridWidget(QWidget):
    """Сетка полей (несколько столбцов), перетаскивание."""

    def __init__(self, parent=None):
        super().__init__(parent)
        lay = QGridLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(2)
        for i, (field, label) in enumerate(AVAILABLE_FIELDS):
            r, c = i // 3, i % 3
            btn = DraggableFieldButton(field, label)
            lay.addWidget(btn, r, c)


# ─────────────────────────── Таблица шаблона ───────────────────────────────

class TemplateGridTable(QTableWidget):
    """Таблица шаблона. Принимает drop полей, перетаскивание из ячеек, объединение, удаление поля, очистка."""

    def __init__(self, rows: int, cols: int, parent=None):
        super().__init__(rows, cols, parent)
        self.setAcceptDrops(True)
        self.setDragEnabled(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.DragDrop)
        self.setDropIndicatorShown(True)
        self.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.setSelectionMode(QAbstractItemView.SelectionMode.ContiguousSelection)
        self._drag_start_cell = None
        self._drag_start_pos = None
        self.horizontalHeader().setMinimumSectionSize(90)
        self.verticalHeader().setMinimumSectionSize(44)
        self.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.verticalHeader().setVisible(True)
        h_labels = [chr(65 + i) for i in range(cols)]
        self.setHorizontalHeaderLabels(h_labels)
        v_labels = []
        for i in range(rows):
            if i < 3:
                v_labels.append(f"Заголовок {i + 1}")
            else:
                v_labels.append(f"Данные {i - 2}")
        self.setVerticalHeaderLabels(v_labels)
        for r in range(rows):
            for c in range(cols):
                it = QTableWidgetItem("")
                it.setData(Qt.ItemDataRole.UserRole, None)
                it.setFlags(it.flags() | Qt.ItemFlag.ItemIsDragEnabled)
                self.setItem(r, c, it)

    def _get_item_at(self, row: int, col: int) -> QTableWidgetItem | None:
        """Возвращает виджет-владелец ячейки (при объединении — верхняя левая ячейка span)."""
        it = self.item(row, col)
        if it is not None:
            return it
        for r0 in range(self.rowCount()):
            for c0 in range(self.columnCount()):
                it0 = self.item(r0, c0)
                if it0 is None:
                    continue
                rs, cs = self.rowSpan(r0, c0), self.columnSpan(r0, c0)
                if r0 <= row < r0 + rs and c0 <= col < c0 + cs:
                    return it0
        return None

    def mousePressEvent(self, event):
        self._drag_start_cell = None
        self._drag_start_pos = None
        idx = self.indexAt(event.position().toPoint() if hasattr(event, "position") else event.pos())
        if idx.isValid():
            self._drag_start_cell = (idx.row(), idx.column())
            self._drag_start_pos = event.position().toPoint() if hasattr(event, "position") else event.pos()
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        """Инициирует перетаскивание из ячейки при движении мыши (надёжнее, чем полагаться на startDrag)."""
        if self._drag_start_cell is not None and self._drag_start_pos is not None:
            pos = event.position().toPoint() if hasattr(event, "position") else event.pos()
            if (pos - self._drag_start_pos).manhattanLength() > 8:
                r, c = self._drag_start_cell
                item = self._get_item_at(r, c)
                if item is not None:
                    field = item.data(Qt.ItemDataRole.UserRole)
                    text = (item.text() or "").strip()
                    label = text or (FIELD_LABEL_MAP.get(field, field) if field else "")
                    if field or label:
                        src_r, src_c = self._get_cell_of_item(item)
                        md = QMimeData()
                        payload = f"{field or ''}||{label}||{src_r},{src_c}"
                        md.setData(MIME_FIELD_FROM_CELL, payload.encode("utf-8"))
                        md.setData(MIME_FIELD, f"{field or ''}||{label}".encode("utf-8"))
                        md.setText(f"{field or ''}||{label}")
                        drag = QDrag(self)
                        drag.setMimeData(md)
                        rect = self.visualItemRect(item)
                        pixmap = QPixmap(rect.size())
                        pixmap.fill(Qt.GlobalColor.transparent)
                        self.viewport().render(pixmap, QPoint(), rect)
                        drag.setPixmap(pixmap)
                        drag.setHotSpot(QPoint(pixmap.width() // 2, pixmap.height() // 2))
                        drag.exec(Qt.DropAction.MoveAction)
                        self._drag_start_cell = None
                        self._drag_start_pos = None
                        return
        self._drag_start_cell = None
        self._drag_start_pos = None
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        self._drag_start_cell = None
        self._drag_start_pos = None
        super().mouseReleaseEvent(event)

    def startDrag(self, supportedActions):
        if self._drag_start_cell is not None:
            r, c = self._drag_start_cell
            item = self._get_item_at(r, c)
            if item is not None:
                field = item.data(Qt.ItemDataRole.UserRole)
                text = (item.text() or "").strip()
                label = text or (FIELD_LABEL_MAP.get(field, field) if field else "")
                if field or label:
                    src_r, src_c = self._get_cell_of_item(item)  # верхняя левая для объединённых
                    md = QMimeData()
                    # Перетаскивание из ячейки — move (очистить источник после drop). Формат: field||label||r,c
                    payload = f"{field or ''}||{label}||{src_r},{src_c}"
                    md.setData(MIME_FIELD_FROM_CELL, payload.encode("utf-8"))
                    md.setData(MIME_FIELD, f"{field or ''}||{label}".encode("utf-8"))
                    md.setText(f"{field or ''}||{label}")
                    drag = QDrag(self)
                    drag.setMimeData(md)
                    rect = self.visualItemRect(item)
                    pixmap = QPixmap(rect.size())
                    pixmap.fill(Qt.GlobalColor.transparent)
                    self.viewport().render(pixmap, QPoint(), rect)
                    drag.setPixmap(pixmap)
                    drag.setHotSpot(QPoint(pixmap.width() // 2, pixmap.height() // 2))
                    drag.exec(Qt.DropAction.MoveAction)
                    self._drag_start_cell = None
                    return
        self._drag_start_cell = None

    def dragEnterEvent(self, event):
        if (event.mimeData().hasFormat(MIME_FIELD) or event.mimeData().hasFormat(MIME_FIELD_FROM_CELL)
                or event.mimeData().hasText()):
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        if (event.mimeData().hasFormat(MIME_FIELD) or event.mimeData().hasFormat(MIME_FIELD_FROM_CELL)
                or event.mimeData().hasText()):
            event.acceptProposedAction()

    def dropEvent(self, event):
        md = event.mimeData()
        source_cell = None  # (r, c) если перетаскивание из ячейки (move)
        if md.hasFormat(MIME_FIELD_FROM_CELL):
            raw = md.data(MIME_FIELD_FROM_CELL).data().decode("utf-8")
            parts = raw.split("||", 2)  # ["field", "label", "r,c"]
            if len(parts) >= 3:
                try:
                    rc_str = parts[2]
                    src_r, src_c = map(int, rc_str.split(",", 1))
                    source_cell = (src_r, src_c)
                except (ValueError, IndexError):
                    pass
            text = parts[0] + "||" + parts[1] if len(parts) >= 2 else raw
        else:
            text = md.text()
        if "||" in text:
            field, label = text.split("||", 1)
        else:
            field, label = "", text

        # Целевые ячейки: либо выделение (все выбранные), либо одна ячейка под курсором
        targets = set()
        sel = self.selectedRanges()
        pos = event.position().toPoint()
        cell = self.indexAt(pos)

        if sel and cell.isValid():
            r_drop, c_drop = cell.row(), cell.column()
            in_selection = any(
                rng.topRow() <= r_drop <= rng.bottomRow() and rng.leftColumn() <= c_drop <= rng.rightColumn()
                for rng in sel
            )
            if in_selection:
                for rng in sel:
                    for r in range(rng.topRow(), rng.bottomRow() + 1):
                        for c in range(rng.leftColumn(), rng.rightColumn() + 1):
                            it = self._get_item_at(r, c)
                            if it is not None:
                                rc = self._get_cell_of_item(it)
                                targets.add(rc)
            else:
                it = self._get_item_at(r_drop, c_drop)
                if it is not None:
                    targets.add(self._get_cell_of_item(it))
        elif cell.isValid():
            it = self._get_item_at(cell.row(), cell.column())
            if it is not None:
                targets.add(self._get_cell_of_item(it))

        for (r, c) in targets:
            item = self.item(r, c)
            if item is not None:
                item.setText(label)
                item.setData(Qt.ItemDataRole.UserRole, field if field else None)
                item.setToolTip(FIELD_DESCRIPTIONS.get(field, "") if field else "")

        # Перемещение из ячейки: очистить источник, если цель отличается
        if source_cell is not None:
            src_r, src_c = source_cell
            src_in_targets = any(
                (src_r, src_c) == self._get_cell_of_item(self._get_item_at(tr, tc))
                for (tr, tc) in targets
            )
            if not src_in_targets:
                src_item = self.item(src_r, src_c)
                if src_item is not None:
                    src_item.setText("")
                    src_item.setData(Qt.ItemDataRole.UserRole, None)
                    src_item.setToolTip("")
        event.acceptProposedAction()

    def _get_visible_item(self, row: int, col: int) -> QTableWidgetItem | None:
        """Возвращает виджет-владелец ячейки (при объединении — ячейка верхнего левого угла)."""
        return self._get_item_at(row, col)

    def _get_cell_of_item(self, item: QTableWidgetItem) -> tuple[int, int]:
        for r in range(self.rowCount()):
            for c in range(self.columnCount()):
                if self.item(r, c) is item:
                    return (r, c)
        return (0, 0)

    def get_grid_and_merges(self) -> tuple[list, list]:
        """Собирает из таблицы grid и merges [(r,c,rowSpan,colSpan)]."""
        grid = []
        merges = []
        rows, cols = self.rowCount(), self.columnCount()
        for r in range(rows):
            row = []
            for c in range(cols):
                item = self.item(r, c)
                if item is None:
                    row.append({"text": "", "field": None})
                    continue
                rs = self.rowSpan(r, c)
                cs = self.columnSpan(r, c)
                if rs > 1 or cs > 1:
                    merges.append((r, c, rs, cs))
                row.append({
                    "text": (item.text() or "").strip(),
                    "field": item.data(Qt.ItemDataRole.UserRole),
                })
            grid.append(row)
        return grid, merges


# ─────────────────────────── Диалог редактора шаблона ─────────────────────

def _grid_from_columns(columns: list, rows: int = None, cols: int = None) -> tuple[list, list]:
    """Строит сетку из старого формата columns (для миграции). rows/cols — размер таблицы."""
    r_total = rows or data_store.GRID_ROWS
    c_total = cols or data_store.GRID_COLS
    grid = []
    for r in range(r_total):
        row = []
        for c in range(c_total):
            if r == 0 and c < len(columns):
                col = columns[c]
                label = data_store.get_column_label(col)
                row.append({"text": label, "field": col.get("field")})
            else:
                row.append({"text": "", "field": None})
        grid.append(row)
    return grid, []


class TemplateEditorDialog(QDialog):
    """Диалог редактирования шаблона: таблица с настраиваемым размером, строка заголовка, перетаскивание полей."""

    def __init__(self, template: dict, unique_products: list, parent=None):
        super().__init__(parent)
        self._tmpl = copy.deepcopy(template)
        self._unique_products = unique_products
        self._grid_rows = self._tmpl.get("gridRows", data_store.GRID_ROWS)
        self._grid_cols = self._tmpl.get("gridCols", data_store.GRID_COLS)
        self._grid_rows = max(4, min(20, self._grid_rows))
        self._grid_cols = max(4, min(16, self._grid_cols))
        self.setWindowTitle(f"Редактор шаблона: {template['name']}")
        self.setMinimumSize(1200, 720)
        self.resize(1280, 760)
        self.setModal(True)
        self._build_ui()
        self._load_grid()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(12)

        # Название
        name_row = QHBoxLayout()
        name_row.addWidget(QLabel("Название шаблона:"))
        self.le_name = CommitLineEdit(self._tmpl["name"])
        self.le_name.commit.connect(self._on_name_commit)
        name_row.addWidget(self.le_name)
        root.addLayout(name_row)

        # Формат и привязка к отделу
        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Формат файла:"))
        self.combo_format = QComboBox()
        self.combo_format.addItem("Столбцы (сетка)", "")
        self.combo_format.addItem("Широкий (Wide)", "wide")
        self.combo_format.addItem("Строчный (Rows)", "rows")
        current_fmt = self._tmpl.get("format", "")
        for i in range(self.combo_format.count()):
            if self.combo_format.itemData(i) == current_fmt:
                self.combo_format.setCurrentIndex(i)
                break
        row2.addWidget(self.combo_format)
        row2.addStretch()
        root.addLayout(row2)

        # Привязка к отделам/подотделам (множественный выбор)
        dept_row = QHBoxLayout()
        dept_row.addWidget(QLabel("Привязать к отделам:"))
        dept_hint = QLabel("(пусто = шаблон по умолчанию)")
        dept_hint.setObjectName("hintLabel")
        dept_row.addWidget(dept_hint)
        dept_row.addStretch()
        root.addLayout(dept_row)
        dept_scroll = QScrollArea()
        dept_scroll.setWidgetResizable(True)
        dept_scroll.setFrameShape(QFrame.Shape.NoFrame)
        dept_scroll.setMaximumHeight(120)
        dept_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        dept_container = QWidget()
        self.dept_checkboxes_lay = QVBoxLayout(dept_container)
        self.dept_checkboxes_lay.setContentsMargins(0, 4, 0, 4)
        self._dept_checkboxes: list[tuple[str, QCheckBox]] = []
        for key, name in data_store.get_department_choices():
            if not key:
                continue
            cb = QCheckBox(name)
            cb.stateChanged.connect(self._apply_title_row)
            self._dept_checkboxes.append((key, cb))
            self.dept_checkboxes_lay.addWidget(cb)
        dept_scroll.setWidget(dept_container)
        root.addWidget(dept_scroll)

        # Таблица и поля — выше, чтобы удобнее работать
        splitter = QSplitter(Qt.Orientation.Horizontal)
        left = QWidget()
        left.setMaximumWidth(280)
        left_lay = QVBoxLayout(left)
        left_lay.setContentsMargins(0, 0, 0, 0)
        left_lay.addWidget(QLabel("Поля (перетащите в ячейки)"))
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setWidget(FieldsGridWidget())
        left_lay.addWidget(scroll)
        splitter.addWidget(left)

        right = QWidget()
        right_lay = QVBoxLayout(right)
        right_lay.setContentsMargins(0, 0, 0, 0)
        table_header = QHBoxLayout()
        table_header.addWidget(QLabel("Таблица шаблона"))
        table_header.addStretch()
        btn_clear_table = QPushButton("Очистить таблицу")
        btn_clear_table.setObjectName("btnSecondary")
        btn_clear_table.setToolTip("Удалить поля из всех ячеек и разъединить объединённые ячейки")
        btn_clear_table.clicked.connect(self._on_clear_table)
        table_header.addWidget(btn_clear_table)
        right_lay.addLayout(table_header)
        self.table = TemplateGridTable(self._grid_rows, self._grid_cols, self)
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._on_table_context_menu)
        right_lay.addWidget(self.table)
        splitter.addWidget(right)
        splitter.setSizes([260, 900])
        root.addWidget(splitter)

        # Строка 1 (заголовок): авто, название отдела, дата. Тип (основной/увеличение) подставится при создании файла.
        title_row = self._tmpl.get("titleRow") or {}
        row_title = QHBoxLayout()
        row_title.addWidget(QLabel("Строка 1 (заголовок):"))
        row_title.addWidget(QLabel("Авто"))
        self.chk_title_auto = ToggleSwitch()
        self.chk_title_auto.setChecked(title_row.get("auto", True))
        self.chk_title_auto.stateChanged.connect(self._apply_title_row)
        row_title.addWidget(self.chk_title_auto)
        row_title.addWidget(QLabel("Название отдела"))
        self.chk_title_dept = ToggleSwitch()
        self.chk_title_dept.setChecked(title_row.get("includeDept", True))
        self.chk_title_dept.stateChanged.connect(self._apply_title_row)
        row_title.addWidget(self.chk_title_dept)
        row_title.addWidget(QLabel("Дата:"))
        self.date_title = QDateEdit()
        self.date_title.setCalendarPopup(True)
        self.date_title.setDate(_parse_date_to_qdate(title_row.get("date", "")))
        self.date_title.dateChanged.connect(self._apply_title_row)
        row_title.addWidget(self.date_title)
        lbl_type_hint = QLabel("(основной/увеличение — по типу принятого файла)")
        lbl_type_hint.setObjectName("hintLabel")
        row_title.addWidget(lbl_type_hint)
        row_title.addStretch()
        root.addLayout(row_title)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_cancel = QPushButton("Отмена")
        btn_cancel.setObjectName("btnSecondary")
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_cancel)
        btn_save = QPushButton("Сохранить")
        btn_save.setObjectName("btnPrimary")
        btn_save.clicked.connect(self._on_save)
        btn_row.addWidget(btn_save)
        root.addLayout(btn_row)

    def _on_name_commit(self):
        t = self.le_name.text().strip()
        if t:
            self._tmpl["name"] = t

    def _get_selected_dept_keys(self) -> list[str]:
        """Возвращает список выбранных ключей отделов."""
        return [key for key, cb in self._dept_checkboxes if cb.isChecked()]

    def _get_first_dept_display_name(self) -> str:
        """Имя первого выбранного отдела для заголовка."""
        keys = self._get_selected_dept_keys()
        if not keys:
            return ""
        return data_store.get_department_display_name(keys[0])

    def _apply_title_row(self):
        """Записывает в ячейку (0,0) заголовок по настройкам авто-строки."""
        if not self.chk_title_auto.isChecked():
            return
        dept_name = self._get_first_dept_display_name()
        date_str = self.date_title.date().toString("dd.MM.yyyy")
        title = _build_title_string(
            self.chk_title_dept.isChecked(), dept_name, date_str, None
        )
        if self.table.rowCount() > 0 and self.table.columnCount() > 0:
            it = self.table.item(0, 0)
            if it:
                it.setText(title)

    def _load_grid(self):
        # Загрузка выбранных отделов
        dept_keys = set(self._tmpl.get("deptKeys") or ([self._tmpl.get("deptKey")] if self._tmpl.get("deptKey") else []))
        for key, cb in self._dept_checkboxes:
            cb.blockSignals(True)
            cb.setChecked(key in dept_keys)
            cb.blockSignals(False)

        grid = self._tmpl.get("grid")
        merges = self._tmpl.get("merges") or []
        if not grid or len(grid) < self._grid_rows:
            grid, merges = _grid_from_columns(
                self._tmpl.get("columns", []), self._grid_rows, self._grid_cols
            )
        rows = min(len(grid), self.table.rowCount())
        cols = min(len(grid[0]) if grid else 0, self.table.columnCount())
        title_row = self._tmpl.get("titleRow") or {}
        if title_row.get("auto", True) and rows > 0:
            dept_name = self._get_first_dept_display_name() or ""
            if not dept_name and self._tmpl.get("deptKey"):
                dept_name = data_store.get_department_display_name(self._tmpl["deptKey"])
            title = _build_title_string(
                title_row.get("includeDept", True),
                dept_name,
                title_row.get("date", ""),
                None,
            )
            grid[0][0] = {"text": title, "field": grid[0][0].get("field")}
        for r in range(rows):
            for c in range(cols):
                cell = grid[r][c]
                item = self.table.item(r, c)
                if item:
                    item.setText(cell.get("text") or "")
                    f = cell.get("field")
                    item.setData(Qt.ItemDataRole.UserRole, f)
                    item.setToolTip(FIELD_DESCRIPTIONS.get(f, "") if f else "")
        self.table.clearSpans()
        for (r, c, rs, cs) in merges:
            if r < self.table.rowCount() and c < self.table.columnCount() and rs >= 1 and cs >= 1:
                self.table.setSpan(r, c, rs, cs)

    def _on_table_context_menu(self, pos):
        menu = QMenu(self)
        sel = self.table.selectedRanges()
        if sel:
            rng = sel[0]
            top, left = rng.topRow(), rng.leftColumn()
            bottom, right = rng.bottomRow(), rng.rightColumn()
            if top <= bottom and left <= right and (bottom - top > 0 or right - left > 0):
                act_merge = menu.addAction("Объединить ячейки")
            else:
                act_merge = None
        else:
            act_merge = None
        act_unmerge = menu.addAction("Разъединить ячейки")
        menu.addSeparator()
        n_cells = self._count_selected_cell_owners()
        if n_cells > 0:
            act_remove_field = menu.addAction(
                "Удалить поле из ячейки" if n_cells == 1 else f"Удалить поле из выбранных ячеек ({n_cells})"
            )
        else:
            act_remove_field = None
        act_clear_all = menu.addAction("Очистить таблицу")
        action = menu.exec(self.table.viewport().mapToGlobal(pos))
        if action == act_merge and sel:
            rng = sel[0]
            self.table.setSpan(rng.topRow(), rng.leftColumn(),
                              rng.bottomRow() - rng.topRow() + 1,
                              rng.rightColumn() - rng.leftColumn() + 1)
        elif action == act_unmerge:
            r, c = self.table.currentRow(), self.table.currentColumn()
            if r >= 0 and c >= 0:
                rs, cs = self.table.rowSpan(r, c), self.table.columnSpan(r, c)
                self.table.setSpan(r, c, 1, 1)
                for rr in range(r, min(r + rs, self.table.rowCount())):
                    for cc in range(c, min(c + cs, self.table.columnCount())):
                        if (rr, cc) != (r, c) and self.table.item(rr, cc) is None:
                            it = QTableWidgetItem("")
                            it.setData(Qt.ItemDataRole.UserRole, None)
                            it.setFlags(it.flags() | Qt.ItemFlag.ItemIsDragEnabled)
                            self.table.setItem(rr, cc, it)
        elif action == act_remove_field:
            self._remove_field_from_selection()
        elif action == act_clear_all:
            self._on_clear_table()

    def _count_selected_cell_owners(self) -> int:
        """Количество «владеющих» ячеек в выделении (для объединённых — одна на span)."""
        targets = set()
        for rng in self.table.selectedRanges():
            for r in range(rng.topRow(), rng.bottomRow() + 1):
                for c in range(rng.leftColumn(), rng.rightColumn() + 1):
                    it = self.table._get_item_at(r, c)
                    if it is not None:
                        rc = self.table._get_cell_of_item(it)
                        targets.add(rc)
        return len(targets)

    def _remove_field_from_selection(self):
        """Очищает поле и текст в выбранных ячейках (учитывая объединённые — по одной на span)."""
        targets = set()
        for rng in self.table.selectedRanges():
            for r in range(rng.topRow(), rng.bottomRow() + 1):
                for c in range(rng.leftColumn(), rng.rightColumn() + 1):
                    it = self.table._get_item_at(r, c)
                    if it is not None:
                        targets.add(self.table._get_cell_of_item(it))
        for (r, c) in targets:
            it = self.table.item(r, c)
            if it is not None:
                it.setText("")
                it.setData(Qt.ItemDataRole.UserRole, None)
                it.setToolTip("")

    def _on_clear_table(self):
        """Очищает все ячейки и разъединяет объединения."""
        reply = QMessageBox.question(
            self, "Очистить таблицу",
            "Удалить поля из всех ячеек и разъединить объединённые ячейки?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        self.table.clearSpans()
        for r in range(self.table.rowCount()):
            for c in range(self.table.columnCount()):
                it = self.table.item(r, c)
                if it is not None:
                    it.setText("")
                    it.setData(Qt.ItemDataRole.UserRole, None)
                    it.setToolTip("")
                else:
                    it = QTableWidgetItem("")
                    it.setData(Qt.ItemDataRole.UserRole, None)
                    it.setFlags(it.flags() | Qt.ItemFlag.ItemIsDragEnabled)
                    self.table.setItem(r, c, it)
        self._apply_title_row()

    def _on_save(self):
        name = self.le_name.text().strip()
        if not name:
            QMessageBox.warning(self, "Ошибка", "Введите название шаблона.")
            return
        self._tmpl["name"] = name
        self._tmpl["format"] = self.combo_format.currentData() or ""
        dept_keys = self._get_selected_dept_keys()
        grid, merges = self.table.get_grid_and_merges()
        title_row = {
            "auto": self.chk_title_auto.isChecked(),
            "includeDept": self.chk_title_dept.isChecked(),
            "date": self.date_title.date().toString("dd.MM.yyyy"),
        }
        data_store.save_template(
            self._tmpl["id"], name,
            [],
            dept_keys=dept_keys,
            fmt=self._tmpl["format"],
            grid=grid,
            merges=merges,
            grid_rows=self.table.rowCount(),
            grid_cols=self.table.columnCount(),
            title_row=title_row,
        )
        self.accept()


# ─────────────────────────── Главный диалог шаблонов ──────────────────────

def _template_preview_html(tmpl: dict) -> str:
    """Генерирует HTML превью шаблона с учётом объединённых ячеек."""
    grid = tmpl.get("grid")
    merges = tmpl.get("merges") or []
    rows_tmpl = tmpl.get("gridRows", data_store.GRID_ROWS)
    cols_tmpl = tmpl.get("gridCols", data_store.GRID_COLS)
    if not grid or len(grid) < rows_tmpl:
        grid, _ = _grid_from_columns(tmpl.get("columns", []), rows_tmpl, cols_tmpl)

    # Карта: (r,c) -> (rowspan, colspan) для верхней левой ячейки объединения
    merge_map: dict[tuple[int, int], tuple[int, int]] = {}
    for (r0, c0, rs, cs) in merges:
        if rs > 0 and cs > 0:
            merge_map[(r0, c0)] = (rs, cs)

    # Множество ячеек, входящих в объединение (не верхняя левая)
    covered: set[tuple[int, int]] = set()
    for (r0, c0), (rs, cs) in merge_map.items():
        for rr in range(r0, min(r0 + rs, rows_tmpl)):
            for cc in range(c0, min(c0 + cs, cols_tmpl)):
                if (rr, cc) != (r0, c0):
                    covered.add((rr, cc))

    cell_style = "border:1px solid #888; padding:8px 12px; font-size:12px;"
    header_style = "border:2px solid #333; padding:8px 12px; font-size:12px; font-weight:bold; background:#f0f0f0;"
    rows_html = []
    for r in range(min(len(grid), rows_tmpl)):
        cells = []
        for c in range(min(len(grid[r]) if grid[r] else 0, cols_tmpl)):
            if (r, c) in covered:
                continue
            cell = grid[r][c]
            text = (cell.get("text") or "").strip() or "—"
            tag = "th" if r == 0 else "td"
            style = header_style if r == 0 else cell_style
            rowspan = ""
            colspan = ""
            if (r, c) in merge_map:
                rs, cs = merge_map[(r, c)]
                if rs > 1:
                    rowspan = f" rowspan='{rs}'"
                if cs > 1:
                    colspan = f" colspan='{cs}'"
            cells.append(f"<{tag} style='{style}'{rowspan}{colspan}>{text}</{tag}>")
        rows_html.append("<tr>" + "".join(cells) + "</tr>")
    table = "<table style='border-collapse:collapse;'>" + "".join(rows_html) + "</table>"
    dept_keys = tmpl.get("deptKeys") or ([tmpl.get("deptKey")] if tmpl.get("deptKey") else [])
    if dept_keys:
        dept_names = [data_store.get_department_display_name(k) for k in dept_keys]
        dept = ", ".join(dept_names)
    else:
        dept = "Все отделы"
    return f"<html><body style='font-family:Segoe UI,sans-serif;'><p style='margin:0 0 10px 0; font-size:13px;'><b>{tmpl.get('name', '')}</b> · {dept}</p>{table}</body></html>"


def _template_tooltip_html(tmpl: dict) -> str:
    """Генерирует HTML таблицы для тултипа (упрощённый вариант)."""
    return _template_preview_html(tmpl)


class TemplatesDialog(QDialog):
    """Модальный диалог управления шаблонами. Список слева, превью справа (Master-Detail)."""

    def __init__(self, app_state: dict, parent=None):
        super().__init__(parent)
        self.app_state = app_state
        self.setWindowTitle("Шаблоны")
        self.setMinimumSize(900, 520)
        self.setModal(True)
        self._build_ui()
        self._refresh_list()

    def _build_ui(self):
        content = QWidget()
        root = QVBoxLayout(content)
        root.setContentsMargins(20, 16, 20, 16)
        root.setSpacing(12)

        title_row = QHBoxLayout()
        lbl_title = QLabel("Управление шаблонами Excel-файлов")
        lbl_title.setObjectName("sectionTitle")
        title_row.addWidget(lbl_title)
        title_row.addWidget(hint_icon_button(
            self,
            "Шаблон — таблица 5×6 (заголовки и данные). Выберите шаблон — превью справа. Двойной клик — редактор.",
            "Инструкция — Шаблоны\n\n"
            "1. Список слева: выберите шаблон — превью отображается справа.\n"
            "2. Двойной клик — открыть редактор; кнопки Создать, Редактировать, Удалить.\n"
            "3. В редакторе: перетащите поля в ячейки, ПКМ — объединить/разъединить.\n"
            "4. Привязка к отделу — в редакторе.",
            "Инструкция",
        ))
        title_row.addStretch()
        root.addLayout(title_row)

        hint = QLabel(
            "Выберите шаблон — превью отображается справа. Двойной клик — открыть редактор."
        )
        hint.setObjectName("hintLabel")
        hint.setWordWrap(True)
        root.addWidget(hint)

        # Master-Detail: список слева, превью справа
        splitter = QSplitter(Qt.Orientation.Horizontal)

        left_panel = QWidget()
        left_panel.setMinimumWidth(220)
        left_panel.setMaximumWidth(320)
        left_lay = QVBoxLayout(left_panel)
        left_lay.setContentsMargins(0, 0, 0, 0)
        left_lay.addWidget(QLabel("Шаблоны"))
        self.list_templates = QListWidget()
        self.list_templates.setAlternatingRowColors(True)
        self.list_templates.itemDoubleClicked.connect(self._on_edit)
        self.list_templates.itemSelectionChanged.connect(self._on_selection_changed)
        left_lay.addWidget(self.list_templates)
        splitter.addWidget(left_panel)

        right_panel = QGroupBox("Превью шаблона")
        right_lay = QVBoxLayout(right_panel)
        right_lay.setContentsMargins(12, 12, 12, 12)
        self.preview_browser = QTextBrowser()
        self.preview_browser.setMinimumHeight(280)
        self.preview_browser.setOpenExternalLinks(False)
        self.preview_browser.setHtml(
            "<html><body style='font-family:Segoe UI,sans-serif; color:#666;'>"
            "<p style='margin:20px;'>Выберите шаблон в списке слева для отображения превью.</p>"
            "</body></html>"
        )
        right_lay.addWidget(self.preview_browser)
        splitter.addWidget(right_panel)
        splitter.setSizes([260, 520])

        root.addWidget(splitter)

        btn_row = QHBoxLayout()
        btn_new = QPushButton("Создать шаблон")
        btn_new.setObjectName("btnPrimary")
        btn_new.clicked.connect(self._on_create)
        btn_row.addWidget(btn_new)
        btn_edit = QPushButton("Редактировать")
        btn_edit.setObjectName("btnSecondary")
        btn_edit.clicked.connect(self._on_edit)
        btn_row.addWidget(btn_edit)
        btn_del = QPushButton("Удалить")
        btn_del.setObjectName("btnDanger")
        btn_del.clicked.connect(self._on_delete)
        btn_row.addWidget(btn_del)
        btn_row.addStretch()
        btn_close = QPushButton("Закрыть")
        btn_close.setObjectName("btnSecondary")
        btn_close.clicked.connect(self.accept)
        btn_row.addWidget(btn_close)
        root.addLayout(btn_row)

        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setWidget(content)
        main_lay = QVBoxLayout(self)
        main_lay.setContentsMargins(0, 0, 0, 0)
        main_lay.addWidget(scroll)

    def _refresh_list(self):
        self.list_templates.blockSignals(True)
        self.list_templates.clear()
        templates = data_store.get_ref("templates") or []
        for tmpl in templates:
            dept_keys = tmpl.get("deptKeys") or ([tmpl.get("deptKey")] if tmpl.get("deptKey") else [])
            if dept_keys:
                dept_str = ", ".join(data_store.get_department_display_name(k) for k in dept_keys[:3])
                if len(dept_keys) > 3:
                    dept_str += f" +{len(dept_keys) - 3}"
            else:
                dept_str = "все"
            item = QListWidgetItem(f"{tmpl['name']}  ·  Отделы: {dept_str}")
            item.setData(Qt.ItemDataRole.UserRole, tmpl["id"])
            item.setData(Qt.ItemDataRole.UserRole + 1, tmpl)
            item.setToolTip(_template_tooltip_html(tmpl))
            self.list_templates.addItem(item)
        self.list_templates.blockSignals(False)
        self._on_selection_changed()

    def _on_selection_changed(self):
        """Обновляет панель превью при смене выбранного шаблона."""
        items = self.list_templates.selectedItems()
        if not items:
            self.preview_browser.setHtml(
                "<html><body style='font-family:Segoe UI,sans-serif; color:#666;'>"
                "<p style='margin:20px;'>Выберите шаблон в списке слева для отображения превью.</p>"
                "</body></html>"
            )
            return
        tmpl = items[0].data(Qt.ItemDataRole.UserRole + 1)
        if tmpl:
            self.preview_browser.setHtml(_template_preview_html(tmpl))

    def _get_selected_id(self) -> str | None:
        items = self.list_templates.selectedItems()
        return items[0].data(Qt.ItemDataRole.UserRole) if items else None

    def _on_create(self):
        name, ok = QInputDialog.getText(self, "Новый шаблон", "Введите название нового шаблона:")
        if not ok or not name.strip():
            return
        tmpl = data_store.create_template(name.strip())
        self._refresh_list()
        self._edit_template(tmpl["id"])

    def _on_edit(self):
        tid = self._get_selected_id()
        if not tid:
            QMessageBox.information(self, "Выберите шаблон", "Выберите шаблон из списка для редактирования.")
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
            QMessageBox.information(self, "Выберите шаблон", "Выберите шаблон из списка для удаления.")
            return
        templates = data_store.get_ref("templates") or []
        if len(templates) <= 1:
            QMessageBox.warning(self, "Нельзя удалить", "Должен остаться хотя бы один шаблон.")
            return
        tmpl = next((t for t in templates if t["id"] == tid), None)
        name = tmpl["name"] if tmpl else tid
        reply = QMessageBox.question(
            self, "Удалить шаблон",
            f"Удалить шаблон «{name}»?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            data_store.delete_template(tid)
            self._refresh_list()


def open_modal(parent: QWidget, app_state: dict):
    """Открывает модальный диалог шаблонов."""
    dlg = TemplatesDialog(app_state, parent=parent)
    dlg.exec()
