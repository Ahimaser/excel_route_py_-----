"""
label_template_editor.py — Редактор шаблона этикетки: предпросмотр таблицы и расстановка полей (№ маршрута, дом/строение, количество).
"""
from __future__ import annotations

import os
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QListWidget, QListWidgetItem, QSplitter, QTableWidget, QTableWidgetItem,
    QAbstractItemView, QHeaderView, QMessageBox,
)
from PyQt6.QtCore import Qt, QMimeData

from core import data_store, excel_generator
from ui.styles import STYLESHEET

LABEL_FIELDS = [
    ("routeNumber", "№ маршрута", "12"),
    ("house", "Дом/строение", "д. 5"),
    ("quantity", "Количество", "3.5"),
]
MIME_LABEL_FIELD = "application/x-label-field"


class LabelTemplatePreviewTable(QTableWidget):
    """Таблица предпросмотра: строки шаблона + ячейки для данных. Принимает drop полей."""

    def __init__(self, template_rows: int, ncols: int, matrix: list, parent=None):
        super().__init__(template_rows + 1, ncols, parent)
        self.template_rows = template_rows
        self.ncols = ncols
        self.matrix = matrix
        self.placements: list[dict] = []  # [{"row": r, "col": c, "field": "routeNumber"|"house"|"quantity"}]
        self._sample_by_field = {f[0]: f[2] for f in LABEL_FIELDS}
        self.setAcceptDrops(True)
        self.setDropIndicatorShown(True)
        self.horizontalHeader().setMinimumSectionSize(60)
        self.verticalHeader().setMinimumSectionSize(28)
        self.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.setHorizontalHeaderLabels([chr(65 + i) for i in range(ncols)])
        v_labels = [f"Строка {i + 1}" for i in range(template_rows)] + ["Данные (подстановка)"]
        self.setVerticalHeaderLabels(v_labels)
        self._fill_from_matrix()

    def _fill_from_matrix(self):
        for r in range(self.template_rows):
            for c in range(self.ncols):
                val = ""
                if r < len(self.matrix) and c < len(self.matrix[r]):
                    v = self.matrix[r][c]
                    val = str(v) if v != "" else ""
                it = QTableWidgetItem(val)
                it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable)
                it.setData(Qt.ItemDataRole.UserRole, None)
                self.setItem(r, c, it)
        for c in range(self.ncols):
            it = QTableWidgetItem("")
            it.setData(Qt.ItemDataRole.UserRole, None)
            self.setItem(self.template_rows, c, it)

    def set_placements(self, placements: list[dict]):
        """Устанавливает расстановку полей и обновляет отображение."""
        self.placements = list(placements)
        self._apply_placements_to_cells()

    def _apply_placements_to_cells(self):
        """Очищает ячейки от старых данных полей и заполняет по self.placements."""
        for r in range(self.rowCount()):
            for c in range(self.columnCount()):
                it = self.item(r, c)
                if it is None:
                    continue
                prev_field = it.data(Qt.ItemDataRole.UserRole)
                if prev_field:
                    it.setData(Qt.ItemDataRole.UserRole, None)
                    if r >= self.template_rows:
                        it.setText("")
                    else:
                        # восстановить значение из шаблона
                        if r < len(self.matrix) and c < len(self.matrix[r]):
                            v = self.matrix[r][c]
                            it.setText(str(v) if v != "" else "")
        for pl in self.placements:
            r, c = pl.get("row", 0), pl.get("col", 0)
            field = pl.get("field")
            if field not in self._sample_by_field:
                continue
            if r >= self.rowCount() or c >= self.columnCount():
                continue
            it = self.item(r, c)
            if it is None:
                it = QTableWidgetItem("")
                it.setData(Qt.ItemDataRole.UserRole, None)
                self.setItem(r, c, it)
            it.setText(self._sample_by_field.get(field, ""))
            it.setData(Qt.ItemDataRole.UserRole, field)

    def _cell_at_placement(self, row: int, col: int) -> str | None:
        for pl in self.placements:
            if pl.get("row") == row and pl.get("col") == col:
                return pl.get("field")
        return None

    def dragEnterEvent(self, event):
        if event.mimeData().hasFormat(MIME_LABEL_FIELD) or event.mimeData().hasText():
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        if event.mimeData().hasFormat(MIME_LABEL_FIELD) or event.mimeData().hasText():
            event.acceptProposedAction()

    def dropEvent(self, event):
        pos = event.position().toPoint()
        cell = self.indexAt(pos)
        if not cell.isValid():
            event.ignore()
            return
        r, c = cell.row(), cell.column()
        text = event.mimeData().text()
        parts = text.split("||")
        field = parts[0] if parts else ""
        sample = parts[2] if len(parts) > 2 else self._sample_by_field.get(field, "")
        if field not in ("routeNumber", "house", "quantity"):
            event.acceptProposedAction()
            return
        # убрать предыдущее размещение этого поля
        self.placements = [p for p in self.placements if p.get("field") != field]
        self.placements.append({"row": r, "col": c, "field": field})
        it = self.item(r, c)
        if it is None:
            it = QTableWidgetItem("")
            it.setData(Qt.ItemDataRole.UserRole, None)
            self.setItem(r, c, it)
        it.setText(sample)
        it.setData(Qt.ItemDataRole.UserRole, field)
        event.acceptProposedAction()

    def get_placements(self) -> list[dict]:
        return list(self.placements)


def open_label_template_editor(product_name: str, parent=None) -> bool:
    """
    Открывает диалог редактирования шаблона этикетки для продукта.
    Возвращает True, если шаблон сохранён (layout обновлён).
    """
    products = data_store.get_ref("products") or []
    prod = next((p for p in products if p.get("name") == product_name), None)
    if not prod:
        return False
    template_path = prod.get("labelTemplatePath") or ""
    if not template_path or not os.path.isfile(template_path):
        QMessageBox.warning(parent, "Нет шаблона", "Сначала выберите файл шаблона XLS для этого продукта.")
        return False
    try:
        nrows, ncols, matrix, _last_filled = excel_generator.load_label_template_matrix(template_path)
    except Exception as e:
        QMessageBox.critical(parent, "Ошибка", f"Не удалось загрузить шаблон:\n{e}")
        return False
    # Отображаем все строки и столбцы этикетки из файла шаблона
    template_rows = max(nrows, 1)
    ncols = max(ncols, 1)
    existing = prod.get("labelLayout") or []
    if not isinstance(existing, list):
        existing = []

    dlg = QDialog(parent)
    dlg.setWindowTitle(f"Предпросмотр шаблона этикетки: {product_name}")
    dlg.setMinimumSize(800, 520)
    dlg.resize(900, 560)
    dlg.setStyleSheet(STYLESHEET)
    root = QVBoxLayout(dlg)
    root.setSpacing(12)

    hint = QLabel(
        "Отображаются все строки и столбцы этикетки из файла шаблона. "
        "Перетащите элементы (№ маршрута, дом/строение, количество) в ячейки строки «Данные» или шаблона — так задаётся расстановка для создания файла с этикетками. "
        "Маршруты в файле: по возрастанию номера, сначала все маршруты одного адреса, затем следующего."
    )
    hint.setWordWrap(True)
    hint.setObjectName("stepLabel")
    root.addWidget(hint)

    splitter = QSplitter(Qt.Orientation.Horizontal)
    class _DragList(QListWidget):
        def mimeData(self, items):
            md = QMimeData()
            if items:
                item = items[0]
                fk = item.data(Qt.ItemDataRole.UserRole)
                lb = item.text()
                sm = item.toolTip()
                md.setData(MIME_LABEL_FIELD, f"{fk}||{lb}||{sm}".encode("utf-8"))
                md.setText(f"{fk}||{lb}||{sm}")
            return md

    left_wrap = _DragList()
    left_wrap.setMaximumWidth(180)
    for field_key, label, sample in LABEL_FIELDS:
        it = QListWidgetItem(label)
        it.setData(Qt.ItemDataRole.UserRole, field_key)
        it.setToolTip(sample)
        left_wrap.addItem(it)
    left_wrap.setDragEnabled(True)
    left_wrap.setAcceptDrops(False)
    left_wrap.setDragDropMode(QAbstractItemView.DragDropMode.DragOnly)
    splitter.addWidget(left_wrap)

    table = LabelTemplatePreviewTable(template_rows, ncols, matrix)
    table.set_placements(existing)
    splitter.addWidget(table)
    splitter.setSizes([180, 700])
    root.addWidget(splitter)

    btn_row = QHBoxLayout()
    btn_row.addStretch()
    btn_cancel = QPushButton("Отмена")
    btn_cancel.setObjectName("btnSecondary")
    btn_cancel.clicked.connect(dlg.reject)
    btn_save = QPushButton("Сохранить расстановку")
    btn_save.setObjectName("btnPrimary")
    def on_save():
        placements = table.get_placements()
        data_store.update_product(product_name, labelLayout=placements)
        QMessageBox.information(dlg, "Сохранено", "Расстановка полей сохранена.")
        dlg.accept()
    btn_save.clicked.connect(on_save)
    btn_row.addWidget(btn_cancel)
    btn_row.addWidget(btn_save)
    root.addLayout(btn_row)

    return dlg.exec() == QDialog.DialogCode.Accepted
