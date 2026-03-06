"""
labels_page.py — Страница создания этикеток XLS по шаблонам.
"""
from __future__ import annotations

import os
import logging
from datetime import date, timedelta

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFrame, QMessageBox, QScrollArea, QComboBox, QFileDialog,
    QDialog, QLineEdit, QTableWidget, QTableWidgetItem, QHeaderView,
    QGroupBox, QGridLayout,
)
from PyQt6.QtCore import Qt, pyqtSignal

from core import data_store, excel_generator
from ui.widgets import make_combo_searchable

log = logging.getLogger(__name__)


def _dept_combo_items():
    """Список (display_name, dept_key) для комбо отделов: Все, затем отделы и подотделы."""
    items = [("Все отделы", None)]
    depts = data_store.get_ref("departments") or []
    for dept in sorted((d for d in depts if isinstance(d, dict) and d.get("name")), key=lambda d: (d.get("name") or "").lower()):
        key = dept.get("key")
        if key:
            items.append((dept.get("name", ""), key))
        for sub in sorted((s for s in dept.get("subdepts", []) if isinstance(s, dict) and s.get("name")), key=lambda s: (s.get("name") or "").lower()):
            sk = sub.get("key")
            if sk:
                parent = dept.get("name") or ""
                sub_name = sub.get("name") or ""
                items.append((f"{parent} / {sub_name}" if parent else sub_name, sk))
    return items


class CreateLabelsDialog(QDialog):
    """Диалог создания этикеток: папка, тип (основной/довоз/оба), предпросмотр."""

    def __init__(
        self,
        parent: QWidget,
        routes: list,
        base_dir: str,
        folder_name_default: str,
        scope_product: str | None,
        scope_dept_key: str | None,
        create_for: str,
        preview_only: bool = False,
    ):
        super().__init__(parent)
        self.preview_only = preview_only
        self.setWindowTitle("Предпросмотр этикеток" if preview_only else "Создать этикетки XLS")
        self.routes = routes
        self.base_dir = base_dir or data_store.get_desktop_path()
        self.folder_name_default = folder_name_default
        self.scope_product = scope_product
        self.scope_dept_key = scope_dept_key
        self.create_for = create_for
        self.created_files = []
        self._build_ui()

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setSpacing(16)

        self.folder_edit = None
        self.combo_for = None
        if not self.preview_only:
            gr = QGroupBox("Папка для сохранения")
            gr_lay = QVBoxLayout(gr)
            self.folder_edit = QLineEdit()
            self.folder_edit.setPlaceholderText("Имя папки (можно изменить)")
            self.folder_edit.setText(self.folder_name_default)
            gr_lay.addWidget(self.folder_edit)
            lay.addWidget(gr)

            gr2 = QGroupBox("Создать для")
            gr2_lay = QVBoxLayout(gr2)
            self.combo_for = QComboBox()
            self.combo_for.addItem("Основной", "main")
            self.combo_for.addItem("Довоз", "increase")
            self.combo_for.addItem("Оба", "both")
            idx = self.combo_for.findData(self.create_for)
            if idx >= 0:
                self.combo_for.setCurrentIndex(idx)
            make_combo_searchable(self.combo_for)
            self.combo_for.currentIndexChanged.connect(self._fill_preview)
            gr2_lay.addWidget(self.combo_for)
            lay.addWidget(gr2)

        gr3 = QGroupBox("Предпросмотр")
        gr3_lay = QVBoxLayout(gr3)
        self.preview_table = QTableWidget()
        self.preview_table.setColumnCount(4)
        self.preview_table.setHorizontalHeaderLabels(["Продукт", "Отдел", "Для", "Маршрутов"])
        self.preview_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.preview_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.preview_table.setColumnHidden(2, True)
        gr3_lay.addWidget(self.preview_table)
        lay.addWidget(gr3)

        btn_lay = QHBoxLayout()
        btn_lay.addStretch()
        if self.preview_only:
            btn_close = QPushButton("Закрыть")
            btn_close.setObjectName("btnSecondary")
            btn_close.clicked.connect(self.accept)
            btn_lay.addWidget(btn_close)
        else:
            btn_cancel = QPushButton("Отмена")
            btn_cancel.setObjectName("btnSecondary")
            btn_cancel.clicked.connect(self.reject)
            btn_create = QPushButton("Создать")
            btn_create.setObjectName("btnPrimary")
            btn_create.clicked.connect(self._on_create)
            btn_lay.addWidget(btn_cancel)
            btn_lay.addWidget(btn_create)
        lay.addLayout(btn_lay)

        self._fill_preview()

    def _current_create_for(self) -> str:
        if self.combo_for:
            return self.combo_for.currentData() or "main"
        return self.create_for

    def _fill_preview(self):
        products_ref = data_store.get_ref("products") or []
        departments_ref = data_store.get_ref("departments") or []
        cf = self._current_create_for()
        show_for_col = cf == "both"
        self.preview_table.setColumnHidden(2, not show_for_col)

        rows: list[tuple[str, str, str, int]] = []
        if cf == "both":
            for ft, label in [("main", "Основной"), ("increase", "Довоз")]:
                preview = excel_generator.labels_preview(
                    self.routes, ft, products_ref, departments_ref,
                    self.scope_product, self.scope_dept_key,
                )
                for prod_name, dept_name, cnt in preview:
                    rows.append((prod_name, dept_name, label, cnt))
        else:
            preview = excel_generator.labels_preview(
                self.routes, cf, products_ref, departments_ref,
                self.scope_product, self.scope_dept_key,
            )
            for prod_name, dept_name, cnt in preview:
                rows.append((prod_name, dept_name, "", cnt))

        self.preview_table.setRowCount(len(rows))
        for r, (p, d, f, c) in enumerate(rows):
            self.preview_table.setItem(r, 0, QTableWidgetItem(p))
            self.preview_table.setItem(r, 1, QTableWidgetItem(d))
            self.preview_table.setItem(r, 2, QTableWidgetItem(f))
            self.preview_table.setItem(r, 3, QTableWidgetItem(str(c)))

    def _on_create(self):
        if not self.folder_edit:
            return
        name = (self.folder_edit.text() or "").strip() or self.folder_name_default
        out_dir = os.path.join(self.base_dir, name) if not os.path.isabs(name) else name
        try:
            os.makedirs(out_dir, exist_ok=True)
        except OSError as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось создать папку: {e}")
            return
        products_ref = data_store.get_ref("products") or []
        departments_ref = data_store.get_ref("departments") or []
        cf = self._current_create_for()
        created: list[str] = []
        try:
            if cf == "both":
                os.makedirs(os.path.join(out_dir, "Основной"), exist_ok=True)
                os.makedirs(os.path.join(out_dir, "Довоз"), exist_ok=True)
                created.extend(excel_generator.generate_labels_from_templates(
                    self.routes, os.path.join(out_dir, "Основной"), "main", products_ref, departments_ref,
                    self.scope_product, self.scope_dept_key,
                ))
                created.extend(excel_generator.generate_labels_from_templates(
                    self.routes, os.path.join(out_dir, "Довоз"), "increase", products_ref, departments_ref,
                    self.scope_product, self.scope_dept_key,
                ))
            else:
                created = excel_generator.generate_labels_from_templates(
                    self.routes, out_dir, cf, products_ref, departments_ref,
                    self.scope_product, self.scope_dept_key,
                )
            self.created_files = created
            if created:
                QMessageBox.information(self, "Готово", f"Создано файлов: {len(created)}\n\n{out_dir}")
                self.accept()
            else:
                QMessageBox.information(self, "Нет файлов", "Нет этикеток для создания.")
        except Exception as e:
            log.exception("create labels")
            QMessageBox.critical(self, "Ошибка", str(e))


class LabelsPage(QWidget):
    """Страница «Этикетки»: создание XLS по шаблонам."""

    go_back = pyqtSignal()
    go_open_routes = pyqtSignal()
    go_process_files = pyqtSignal()

    def __init__(self, app_state: dict):
        super().__init__()
        self.app_state = app_state
        self._build_ui()

    def _build_ui(self):
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
        lay.setSpacing(28)

        h = QHBoxLayout()
        btn_back = QPushButton("← Назад")
        btn_back.setObjectName("btnBack")
        btn_back.clicked.connect(self.go_back.emit)
        h.addWidget(btn_back)
        title = QLabel("Этикетки")
        title.setObjectName("sectionTitle")
        h.addWidget(title)
        h.addStretch()
        try:
            from ui.pages.labels_settings_dialog import open_labels_settings_dialog
            btn_settings = QPushButton("Настройки этикеток")
            btn_settings.setObjectName("btnSecondary")
            btn_settings.clicked.connect(lambda: open_labels_settings_dialog(self))
            h.addWidget(btn_settings)
        except Exception:
            pass
        lay.addLayout(h)

        hint = QLabel(
            "Создайте этикетки из загруженных маршрутов по шаблонам XLS. "
            "Выберите продукт и загрузите шаблон — в шаблон после последней заполненной строки "
            "добавляется строка с № маршрута, дом/строение и количеством; на одном листе — копии для всех маршрутов с разрывом страницы."
        )
        hint.setObjectName("stepLabel")
        hint.setWordWrap(True)
        lay.addWidget(hint)

        # Строка загрузки шаблона: продукт + кнопка + текущий файл
        template_row = QFrame()
        template_row.setObjectName("card")
        tr_lay = QVBoxLayout(template_row)
        tr_lay.setContentsMargins(16, 12, 16, 12)
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("Продукт:"))
        self.combo_label_product = QComboBox()
        self.combo_label_product.setMinimumWidth(220)
        self._fill_label_products_combo()
        make_combo_searchable(self.combo_label_product)
        self.combo_label_product.currentIndexChanged.connect(self._on_label_product_changed)
        row1.addWidget(self.combo_label_product)
        btn_load_tpl = QPushButton("Загрузить шаблон")
        btn_load_tpl.setObjectName("btnSecondary")
        btn_load_tpl.clicked.connect(self._on_load_label_template)
        row1.addWidget(btn_load_tpl)
        self.lbl_current_template = QLabel("—")
        self.lbl_current_template.setObjectName("stepLabel")
        self.lbl_current_template.setWordWrap(True)
        row1.addWidget(self.lbl_current_template)
        row1.addStretch()
        tr_lay.addLayout(row1)
        lay.addWidget(template_row)

        # Область генерации: для кого, что создавать
        scope_row = QFrame()
        scope_row.setObjectName("card")
        scope_lay = QGridLayout(scope_row)
        scope_lay.setContentsMargins(16, 12, 16, 12)
        scope_lay.addWidget(QLabel("Создать:"), 0, 0)
        self.combo_scope = QComboBox()
        self.combo_scope.addItem("Все продукты с шаблонами", "all")
        self.combo_scope.addItem("Только выбранный продукт", "product")
        self.combo_scope.addItem("Только выбранный отдел", "department")
        make_combo_searchable(self.combo_scope)
        self.combo_scope.currentIndexChanged.connect(self._on_scope_changed)
        scope_lay.addWidget(self.combo_scope, 0, 1)
        scope_lay.addWidget(QLabel("Отдел:"), 1, 0)
        self.combo_department = QComboBox()
        for disp, key in _dept_combo_items():
            self.combo_department.addItem(disp, key)
        self.combo_department.setMinimumWidth(220)
        make_combo_searchable(self.combo_department)
        scope_lay.addWidget(self.combo_department, 1, 1)
        scope_lay.addWidget(QLabel("Создать для:"), 2, 0)
        self.combo_create_for = QComboBox()
        self.combo_create_for.addItem("Основной", "main")
        self.combo_create_for.addItem("Довоз", "increase")
        self.combo_create_for.addItem("Оба (отдельные файлы)", "both")
        idx_both = self.combo_create_for.findData("both")
        if idx_both >= 0:
            self.combo_create_for.setCurrentIndex(idx_both)
        make_combo_searchable(self.combo_create_for)
        scope_lay.addWidget(self.combo_create_for, 2, 1)
        self._on_scope_changed()
        lay.addWidget(scope_row)

        self.no_data_frame = QFrame()
        self.no_data_frame.setObjectName("card")
        no_lay = QVBoxLayout(self.no_data_frame)
        no_data_lbl = QLabel("Маршруты не загружены. Откройте последние или обработайте новые файлы.")
        no_data_lbl.setObjectName("stepLabel")
        no_data_lbl.setWordWrap(True)
        no_lay.addWidget(no_data_lbl)
        btn_row = QHBoxLayout()
        btn_open = QPushButton("Открыть последние маршруты")
        btn_open.setObjectName("btnPrimary")
        btn_open.clicked.connect(self.go_open_routes.emit)
        btn_process = QPushButton("Обработать файлы")
        btn_process.setObjectName("btnSecondary")
        btn_process.clicked.connect(self.go_process_files.emit)
        btn_row.addWidget(btn_open)
        btn_row.addWidget(btn_process)
        btn_row.addStretch()
        no_lay.addLayout(btn_row)
        lay.addWidget(self.no_data_frame)

        self.cards_frame = QFrame()
        cards_lay = QVBoxLayout(self.cards_frame)
        lbl_cards = QLabel("Этикетки из шаблонов (XLS)")
        lbl_cards.setObjectName("subsectionLabel")
        cards_lay.addWidget(lbl_cards)
        btn_row = QHBoxLayout()
        self.btn_preview = QPushButton("Предпросмотр")
        self.btn_preview.setObjectName("btnSecondary")
        self.btn_preview.clicked.connect(self._on_preview)
        self.btn_xls = QPushButton("Создать XLS по шаблонам")
        self.btn_xls.setObjectName("btnPrimary")
        self.btn_xls.clicked.connect(self._on_labels_from_templates)
        btn_row.addWidget(self.btn_preview)
        btn_row.addWidget(self.btn_xls)
        btn_row.addStretch()
        cards_lay.addLayout(btn_row)
        lay.addWidget(self.cards_frame)
        lay.addStretch()

    def _fill_label_products_combo(self):
        products = data_store.get_ref("products") or []
        names = sorted(
            [p.get("name", "") for p in products if p.get("name") and p.get("deptKey")],
            key=str.lower,
        )
        self.combo_label_product.clear()
        self.combo_label_product.addItem("— Выберите продукт —", None)
        for n in names:
            self.combo_label_product.addItem(n, n)

    def _fill_departments_combo(self):
        self.combo_department.clear()
        for disp, key in _dept_combo_items():
            self.combo_department.addItem(disp, key)

    def _on_scope_changed(self):
        scope = self.combo_scope.currentData() or "all"
        self.combo_department.setEnabled(scope == "department")
        self.combo_label_product.setEnabled(scope != "all")

    def _on_label_product_changed(self):
        name = self.combo_label_product.currentData()
        if not name:
            self.lbl_current_template.setText("—")
            return
        products = data_store.get_ref("products") or []
        prod = next((p for p in products if p.get("name") == name), None)
        path = (prod.get("labelTemplatePath") or "").strip() if prod else ""
        self.lbl_current_template.setText(os.path.basename(path) if path else "—")

    def _on_load_label_template(self):
        name = self.combo_label_product.currentData()
        if not name:
            QMessageBox.information(self, "Шаблон", "Сначала выберите продукт.")
            return
        products = data_store.get_ref("products") or []
        prod = next((p for p in products if p.get("name") == name), None)
        start_dir = ""
        if prod and prod.get("labelTemplatePath") and os.path.isfile(prod.get("labelTemplatePath", "")):
            start_dir = os.path.dirname(prod["labelTemplatePath"])
        path, _ = QFileDialog.getOpenFileName(
            self, "Шаблон этикетки (XLS)",
            start_dir or None,
            "Excel 97-2003 (*.xls);;Все файлы (*)",
        )
        if path:
            path = os.path.normpath(path)
            if path.lower().endswith(".xlsx"):
                QMessageBox.warning(
                    self, "Формат файла",
                    "Поддерживается только формат Excel 97-2003 (.xls). Выберите файл .xls.",
                )
                return
            if data_store.set_product_label_template(name, path):
                self.lbl_current_template.setText(os.path.basename(path))
            else:
                QMessageBox.warning(
                    self, "Ошибка",
                    f"Не удалось привязать шаблон к продукту «{name}».",
                )

    def _has_routes(self) -> bool:
        routes = self.app_state.get("filteredRoutes", [])
        return bool([r for r in routes if not r.get("excluded")])

    def refresh(self):
        has = self._has_routes()
        self.no_data_frame.setVisible(not has)
        self.cards_frame.setVisible(has)
        self._fill_label_products_combo()
        self._fill_departments_combo()
        self._on_label_product_changed()
        self._on_scope_changed()

    def _get_scope_filters(self) -> tuple[str | None, str | None]:
        scope = self.combo_scope.currentData() or "all"
        if scope == "product":
            return (self.combo_label_product.currentData(), None)
        if scope == "department":
            return (None, self.combo_department.currentData())
        return (None, None)

    def _on_preview(self):
        routes = self.app_state.get("filteredRoutes", [])
        active = [r for r in routes if not r.get("excluded")]
        if not active:
            QMessageBox.warning(self, "Нет данных", "Нет маршрутов для этикеток.")
            return
        only_product, only_dept_key = self._get_scope_filters()
        if self.combo_scope.currentData() == "product" and not only_product:
            QMessageBox.information(self, "Предпросмотр", "Выберите продукт.")
            return
        if self.combo_scope.currentData() == "department" and not only_dept_key:
            QMessageBox.information(self, "Предпросмотр", "Выберите отдел.")
            return
        create_for = self.combo_create_for.currentData() or "main"
        base_dir = self.app_state.get("saveDir") or data_store.get_desktop_path()
        tomorrow = date.today() + timedelta(days=1)
        folder_name = f"Этикетки на {tomorrow:%d.%m.%Y}"
        dlg = CreateLabelsDialog(
            self, active, base_dir, folder_name,
            only_product, only_dept_key, create_for,
            preview_only=True,
        )
        dlg.resize(560, 400)
        dlg.exec()

    def _on_labels_from_templates(self):
        routes = self.app_state.get("filteredRoutes", [])
        active = [r for r in routes if not r.get("excluded")]
        if not active:
            QMessageBox.warning(self, "Нет данных", "Нет маршрутов для этикеток.")
            return
        products_ref = data_store.get_ref("products") or []
        if not any((p.get("labelTemplatePath") or "").strip() and p.get("deptKey") for p in products_ref):
            QMessageBox.information(
                self, "Нет шаблонов",
                "В «Настройках этикеток» укажите шаблон XLS для продуктов с отделом, "
                "либо выберите продукт выше и нажмите «Загрузить шаблон»."
            )
            return
        only_product, only_dept_key = self._get_scope_filters()
        if self.combo_scope.currentData() == "product" and not only_product:
            QMessageBox.information(self, "Создание", "Выберите продукт.")
            return
        if self.combo_scope.currentData() == "department" and not only_dept_key:
            QMessageBox.information(self, "Создание", "Выберите отдел.")
            return
        base_dir = self.app_state.get("saveDir") or data_store.get_desktop_path()
        tomorrow = date.today() + timedelta(days=1)
        folder_name = f"Этикетки на {tomorrow:%d.%m.%Y}"
        create_for = self.combo_create_for.currentData() or "main"
        dlg = CreateLabelsDialog(
            self, active, base_dir, folder_name,
            only_product, only_dept_key, create_for,
        )
        if dlg.exec() == QDialog.DialogCode.Accepted and dlg.created_files:
            if hasattr(self.app_state.get("set_status"), "__call__"):
                self.app_state["set_status"](f"Создано этикеток: {len(dlg.created_files)}")
