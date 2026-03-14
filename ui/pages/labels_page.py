"""
labels_page.py — Страница этикеток в режиме live-preview и печати.
"""
from __future__ import annotations

import os

from PyQt6.QtCore import pyqtSignal, Qt
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
    QFileDialog,
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


def _products_for_dept(dept_key: str) -> list[dict]:
    products = data_store.get_ref("products") or []
    result = []
    for p in products:
        if p.get("deptKey") != dept_key:
            continue
        name = p.get("name")
        if not name:
            continue
        result.append({
            "name": name,
            "unit": p.get("unit") or "",
            "template": p.get("labelTemplatePath") or "",
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
        lay.setContentsMargins(48, 40, 48, 40)
        lay.setSpacing(16)

        row = QHBoxLayout()
        btn_back = QPushButton("← Назад")
        btn_back.setObjectName("btnBack")
        btn_back.clicked.connect(self.go_back.emit)
        row.addWidget(btn_back)
        title = QLabel("Этикетки: live-preview и печать")
        title.setObjectName("sectionTitle")
        row.addWidget(title)
        row.addStretch()
        lay.addLayout(row)

        hint = QLabel(
            "Выберите отдел и продукт, назначьте шаблон .xls, затем откройте предпросмотр "
            "и печатайте без сохранения итоговых файлов."
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
        main_lay.setContentsMargins(16, 14, 16, 14)
        main_lay.setSpacing(10)

        filters = QHBoxLayout()
        filters.addWidget(QLabel("Отдел / подотдел:"))
        self.combo_dept = QComboBox()
        self.combo_dept.currentIndexChanged.connect(self._on_dept_changed)
        filters.addWidget(self.combo_dept, 1)
        filters.addWidget(QLabel("Тип:"))
        self.combo_type = QComboBox()
        self.combo_type.addItem("Основной", "main")
        self.combo_type.addItem("Довоз", "increase")
        filters.addWidget(self.combo_type)
        main_lay.addLayout(filters)

        self.products_list = QListWidget()
        self.products_list.currentItemChanged.connect(self._on_product_changed)
        self.products_list.itemDoubleClicked.connect(self._on_product_double_clicked)
        main_lay.addWidget(self.products_list, 1)

        self.lbl_selected = QLabel("Выберите продукт.")
        self.lbl_selected.setObjectName("hintLabel")
        self.lbl_selected.setWordWrap(True)
        main_lay.addWidget(self.lbl_selected)

        actions = QHBoxLayout()
        self.btn_assign_tpl = QPushButton("Добавить шаблон")
        self.btn_assign_tpl.setObjectName("btnSecondary")
        self.btn_assign_tpl.clicked.connect(self._on_assign_template)
        actions.addWidget(self.btn_assign_tpl)
        self.btn_preview = QPushButton("Предпросмотр")
        self.btn_preview.setObjectName("btnSecondary")
        self.btn_preview.clicked.connect(self._on_preview)
        actions.addWidget(self.btn_preview)
        self.btn_print = QPushButton("Печать этикеток")
        self.btn_print.setObjectName("btnPrimary")
        self.btn_print.clicked.connect(self._on_print)
        actions.addWidget(self.btn_print)
        actions.addStretch()
        main_lay.addLayout(actions)

        lay.addWidget(self.main_frame)
        lay.addStretch()

    def _active_routes(self) -> list[dict]:
        routes = self.app_state.get("filteredRoutes") or self.app_state.get("routes") or []
        return [r for r in routes if not r.get("excluded")]

    def _has_routes(self) -> bool:
        return bool(self._active_routes())

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

    def _on_dept_changed(self) -> None:
        self.products_list.clear()
        dept_key = self.combo_dept.currentData()
        if not dept_key:
            self._update_selected_label()
            return
        for p in _products_for_dept(str(dept_key)):
            text = p["name"]
            if p["template"] and os.path.isfile(p["template"]):
                text += "  [шаблон]"
            item = QListWidgetItem(text)
            item.setData(Qt.ItemDataRole.UserRole, p)
            self.products_list.addItem(item)
        if self.products_list.count():
            self.products_list.setCurrentRow(0)
        self._update_selected_label()

    def _on_product_changed(self, _cur: QListWidgetItem | None, _prev: QListWidgetItem | None) -> None:
        self._update_selected_label()

    def _on_product_double_clicked(self, item: QListWidgetItem) -> None:
        p = item.data(Qt.ItemDataRole.UserRole) or None
        if not p or not (p.get("template") or "").strip() or not os.path.isfile(p.get("template", "")):
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
        self.btn_assign_tpl.setEnabled(has_selection)
        self.btn_preview.setEnabled(has_selection)
        self.btn_print.setEnabled(has_selection)
        if not p:
            self.lbl_selected.setText("Выберите продукт.")
            return
        tpl = p.get("template") or ""
        if tpl and os.path.isfile(tpl):
            self.lbl_selected.setText(f"Продукт: {p['name']}. Шаблон: {os.path.basename(tpl)}")
        else:
            self.lbl_selected.setText(f"Продукт: {p['name']}. Шаблон не назначен.")

    def _on_assign_template(self) -> None:
        p = self._selected_product()
        if not p:
            return
        base_dir = self.app_state.get("saveDir") or data_store.get_desktop_path()
        path, _ = QFileDialog.getOpenFileName(
            self,
            f"Шаблон для продукта «{p['name']}»",
            base_dir,
            "Excel 97-2003 (*.xls);;Все файлы (*)",
        )
        if not path:
            return
        path = os.path.normpath(path)
        if not path.lower().endswith(".xls"):
            QMessageBox.warning(self, "Формат файла", "Поддерживается только шаблон .xls.")
            return
        data_store.update_product(p["name"], labelTemplatePath=path)
        self._on_dept_changed()

    def _build_preview_dialog(self) -> LabelsPrintPreviewDialog | None:
        routes = self._active_routes()
        if not routes:
            QMessageBox.warning(self, "Нет данных", "Нет маршрутов для печати этикеток.")
            return None
        p = self._selected_product()
        if not p:
            QMessageBox.warning(self, "Нет продукта", "Сначала выберите продукт.")
            return None
        tpl = p.get("template") or ""
        if not tpl or not os.path.isfile(tpl):
            QMessageBox.warning(self, "Нет шаблона", "Сначала добавьте шаблон .xls для продукта.")
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
        dlg = self._build_preview_dialog()
        if dlg is not None:
            dlg.exec()

    def _on_print(self) -> None:
        dlg = self._build_preview_dialog()
        if dlg is None:
            return
        dlg.exec()
