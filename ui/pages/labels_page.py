"""
labels_page.py — Страница этикеток в режиме live-preview и печати.
"""
from __future__ import annotations

import os

from PyQt6.QtCore import pyqtSignal, Qt, QTimer
from PyQt6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QFrame,
    QMessageBox,
    QScrollArea,
    QComboBox,
    QListWidget,
    QListWidgetItem,
    QSplitter,
)

from core import data_store
from ui.pages.labels_print_preview_dialog import (
    LabelsPrintPreviewDialog,
    ProductLabelsTableDialog,
)


def _enabled_dept_items() -> list[dict]:
    """Список отделов/подотделов с labelsEnabled=True."""
    result: list[dict] = []
    depts = data_store.get_ref("departments") or []
    for dept in depts:
        dkey = dept.get("key")
        dname = dept.get("name") or dkey
        if dkey and dept.get("labelsEnabled", True):
            result.append({"key": dkey, "name": dname, "display": dname})
        for sub in dept.get("subdepts", []):
            skey = sub.get("key")
            sname = sub.get("name") or skey
            if skey and sub.get("labelsEnabled", True):
                result.append({"key": skey, "name": sname, "display": f"{dname} / {sname}"})
    return result


def _products_for_dept(dept_key: str, routes: list[dict] | None = None) -> list[dict]:
    """
    Продукты отдела для этикеток. Без шаблонов. Если routes заданы — только те, что есть в маршрутах.
    """
    products = data_store.get_ref("products") or []
    products_in_routes: set[str] = set()
    if routes:
        for r in routes:
            if r.get("excluded"):
                continue
            for prod in r.get("products", []):
                name = prod.get("name")
                if name:
                    products_in_routes.add(name)

    result = []
    for p in products:
        if p.get("deptKey") != dept_key:
            continue
        name = p.get("name")
        if not name:
            continue
        if routes is not None and name not in products_in_routes:
            continue
        result.append({
            "name": name,
            "unit": p.get("unit") or "",
            "template": "",  # шаблоны не используются
        })
    return sorted(result, key=lambda x: x["name"].lower())


class LabelsPage(QWidget):
    go_back = pyqtSignal()
    go_open_routes = pyqtSignal()
    go_process_files = pyqtSignal()

    def __init__(self, app_state: dict):
        super().__init__()
        self.app_state = app_state
        self._build_ui()

    def _build_ui(self) -> None:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        content = QWidget()
        scroll.setWidget(content)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(scroll)

        lay = QVBoxLayout(content)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(12)

        row = QHBoxLayout()
        title = QLabel("Этикетки: live-preview и печать")
        title.setObjectName("sectionTitle")
        row.addWidget(title)
        row.addStretch()
        lay.addLayout(row)

        hint = QLabel(
            "Выберите отдел и продукт в списке слева. Файлы этикеток содержат 3 столбца: номер маршрута, дом/строение, количество. "
            "Создаются в папке маршрутов: Основные/этикетки/{отдел}/ или Довоз/этикетки/{отдел}/."
        )
        hint.setObjectName("stepLabel")
        hint.setWordWrap(True)
        lay.addWidget(hint)

        self.no_data_frame = QFrame()
        self.no_data_frame.setObjectName("card")
        nd_lay = QVBoxLayout(self.no_data_frame)
        nd_lay.addWidget(QLabel("Маршруты не загружены. Откройте историю или обработайте новые файлы."))
        row0 = QHBoxLayout()
        btn_open = QPushButton("Открыть историю маршрутов")
        btn_open.setObjectName("btnPrimary")
        btn_open.clicked.connect(self.go_open_routes.emit)
        btn_process = QPushButton("Обработать файлы")
        btn_process.setObjectName("btnSecondary")
        btn_process.clicked.connect(self.go_process_files.emit)
        row0.addWidget(btn_open)
        row0.addWidget(btn_process)
        row0.addStretch()
        nd_lay.addLayout(row0)
        lay.addWidget(self.no_data_frame)

        self.main_frame = QFrame()
        self.main_frame.setObjectName("card")
        main_lay = QVBoxLayout(self.main_frame)
        main_lay.setContentsMargins(16, 12, 16, 12)
        main_lay.setSpacing(10)

        filters = QHBoxLayout()
        filters.addWidget(QLabel("Отдел / подотдел:"))
        self.combo_dept = QComboBox()
        self.combo_dept.setMinimumWidth(200)
        self.combo_dept.currentIndexChanged.connect(self._on_dept_changed)
        filters.addWidget(self.combo_dept, 1)
        filters.addWidget(QLabel("Тип:"))
        self.combo_type = QComboBox()
        self.combo_type.addItem("Основной", "main")
        self.combo_type.addItem("Довоз", "increase")
        self.combo_type.currentIndexChanged.connect(self._on_type_changed)
        filters.addWidget(self.combo_type)
        main_lay.addLayout(filters)

        splitter = QSplitter(Qt.Orientation.Horizontal)

        left_box = QFrame()
        left_lay = QVBoxLayout(left_box)
        left_lay.setContentsMargins(12, 12, 12, 12)
        left_lay.addWidget(QLabel("Продукты"))
        self.products_list = QListWidget()
        self.products_list.setMinimumWidth(220)
        self.products_list.currentItemChanged.connect(self._on_product_changed)
        self.products_list.itemDoubleClicked.connect(self._on_product_double_clicked)
        left_lay.addWidget(self.products_list, 1)
        splitter.addWidget(left_box)

        right_box = QFrame()
        right_lay = QVBoxLayout(right_box)
        right_lay.setContentsMargins(8, 8, 8, 8)
        right_lay.setSpacing(8)
        right_lay.addWidget(QLabel("Действия"))
        self.lbl_selected = QLabel("Выберите продукт.")
        self.lbl_selected.setObjectName("hintLabel")
        self.lbl_selected.setWordWrap(True)
        right_lay.addWidget(self.lbl_selected)

        self.btn_preview = QPushButton("Предпросмотр")
        self.btn_preview.setObjectName("btnSecondary")
        self.btn_preview.clicked.connect(self._on_preview)
        right_lay.addWidget(self.btn_preview)
        self.btn_print = QPushButton("Печать этикеток")
        self.btn_print.setObjectName("btnPrimary")
        self.btn_print.clicked.connect(self._on_print)
        right_lay.addWidget(self.btn_print)
        right_lay.addStretch()
        splitter.addWidget(right_box)

        splitter.setSizes([280, 200])
        main_lay.addWidget(splitter, 1)

        actions_row = QHBoxLayout()
        btn_settings = QPushButton("Настройки этикеток")
        btn_settings.setObjectName("btnSecondary")
        btn_settings.clicked.connect(self._open_labels_settings)
        actions_row.addWidget(btn_settings)
        actions_row.addStretch()
        main_lay.addLayout(actions_row)

        lay.addWidget(self.main_frame)
        lay.addStretch()

    def _active_routes(self) -> list[dict]:
        """Маршруты для выбранного типа (Основной/Довоз)."""
        file_type = self.combo_type.currentData() or "main"
        app_type = self.app_state.get("fileType") or "main"
        routes = self.app_state.get("filteredRoutes") or self.app_state.get("routes") or []
        if app_type != file_type or not routes:
            blob = data_store.get_last_routes(file_type) or {}
            routes = blob.get("filteredRoutes") or blob.get("routes") or []
        return [r for r in routes if not r.get("excluded")]

    def _has_routes(self) -> bool:
        return bool(self._active_routes())

    def showEvent(self, event):
        super().showEvent(event)
        QTimer.singleShot(50, self.refresh)

    def refresh(self) -> None:
        has = self._has_routes()
        self.no_data_frame.setVisible(not has)
        self.main_frame.setVisible(has)
        if not has:
            return
        current_type = str(self.app_state.get("fileType") or "main")
        idx_type = self.combo_type.findData(current_type)
        self.combo_type.setCurrentIndex(idx_type if idx_type >= 0 else 0)
        self._fill_depts()
        if self.combo_dept.count():
            self._on_dept_changed()

    def _fill_depts(self) -> None:
        self.combo_dept.blockSignals(True)
        old_key = self.combo_dept.currentData()
        self.combo_dept.clear()
        items = _enabled_dept_items()
        for it in items:
            self.combo_dept.addItem(it["display"], it["key"])
        self.combo_dept.blockSignals(False)
        if old_key is not None:
            idx = self.combo_dept.findData(old_key)
            if idx >= 0:
                self.combo_dept.setCurrentIndex(idx)
                return
        if self.combo_dept.count():
            self.combo_dept.setCurrentIndex(0)
        else:
            self.products_list.clear()
            self._update_selected_label()

    def _on_type_changed(self) -> None:
        """При смене типа (Основной/Довоз) обновляем список продуктов."""
        self._on_dept_changed()

    def _on_dept_changed(self) -> None:
        self.products_list.clear()
        dept_key = self.combo_dept.currentData()
        if not dept_key:
            self._update_selected_label()
            return
        routes = self._active_routes()
        for p in _products_for_dept(str(dept_key), routes):
            item = QListWidgetItem(p["name"])
            item.setData(Qt.ItemDataRole.UserRole, p)
            self.products_list.addItem(item)
        if self.products_list.count():
            self.products_list.setCurrentRow(0)
        self._update_selected_label()

    def _on_product_changed(self, _cur: QListWidgetItem | None, _prev: QListWidgetItem | None) -> None:
        self._update_selected_label()

    def _on_product_double_clicked(self, item: QListWidgetItem) -> None:
        p = item.data(Qt.ItemDataRole.UserRole) or None
        if not p:
            return
        routes = self._active_routes()
        if not routes:
            QMessageBox.warning(self, "Нет данных", "Нет маршрутов для отображения этикеток.")
            return
        dlg = ProductLabelsTableDialog(
            self,
            routes=routes,
            file_type=self.combo_type.currentData() or "main",
            products_ref=data_store.get_ref("products") or [],
            departments_ref=data_store.get_ref("departments") or [],
            product_name=p["name"],
            dept_key=self.combo_dept.currentData(),
        )
        dlg.exec()

    def _selected_product(self) -> dict | None:
        item = self.products_list.currentItem()
        if not item:
            return None
        return item.data(Qt.ItemDataRole.UserRole) or None

    def _update_selected_label(self) -> None:
        p = self._selected_product()
        has_selection = p is not None
        self.btn_preview.setEnabled(has_selection)
        self.btn_print.setEnabled(has_selection)
        if not p:
            self.lbl_selected.setText("Выберите продукт.")
            return
        self.lbl_selected.setText(f"Продукт: {p['name']}. Файл: 3 столбца (№ маршрута, Дом, Количество).")

    def _open_labels_settings(self) -> None:
        try:
            from ui.pages.labels_settings_dialog import open_labels_settings_dialog
            open_labels_settings_dialog(self)
            self._on_dept_changed()
        except Exception:
            import traceback
            import logging
            logging.getLogger("app").exception("labels_settings")

    def _build_preview_dialog(self) -> LabelsPrintPreviewDialog | None:
        routes = self._active_routes()
        if not routes:
            QMessageBox.warning(self, "Нет данных", "Нет маршрутов для печати этикеток.")
            return None
        p = self._selected_product()
        if not p:
            QMessageBox.warning(self, "Нет продукта", "Сначала выберите продукт.")
            return None
        return LabelsPrintPreviewDialog(
            self,
            routes=routes,
            file_type=self.combo_type.currentData() or "main",
            products_ref=data_store.get_ref("products") or [],
            departments_ref=data_store.get_ref("departments") or [],
            product_name=p["name"],
            dept_key=self.combo_dept.currentData(),
        )

    def _on_preview(self) -> None:
        try:
            dlg = self._build_preview_dialog()
            if dlg is not None:
                dlg.exec()
        except Exception as exc:
            QMessageBox.critical(
                self,
                "Ошибка предпросмотра",
                f"Не удалось открыть предпросмотр этикеток:\n\n{exc}",
            )

    def _on_print(self) -> None:
        try:
            dlg = self._build_preview_dialog()
            if dlg is None:
                return
            dlg.exec()
        except Exception as exc:
            QMessageBox.critical(
                self,
                "Ошибка печати",
                f"Не удалось открыть диалог печати:\n\n{exc}",
            )
