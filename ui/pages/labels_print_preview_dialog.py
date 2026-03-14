from __future__ import annotations

import os
import shutil

from PyQt6.QtCore import QObject, QThread, pyqtSignal, Qt
from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QComboBox,
    QMessageBox,
    QInputDialog,
)

from core import data_store, excel_generator


class _PrepareWorker(QObject):
    finished = pyqtSignal(str, str)  # (xls_path, temp_dir)
    error = pyqtSignal(str)

    def __init__(
        self,
        routes: list[dict],
        file_type: str,
        products_ref: list,
        departments_ref: list,
        product_name: str,
        dept_key: str | None,
    ):
        super().__init__()
        self.routes = routes
        self.file_type = file_type
        self.products_ref = products_ref
        self.departments_ref = departments_ref
        self.product_name = product_name
        self.dept_key = dept_key

    def run(self) -> None:
        try:
            xls_path, temp_dir = excel_generator.prepare_label_temp_file(
                self.routes,
                self.file_type,
                self.products_ref,
                self.departments_ref,
                self.product_name,
                self.dept_key,
            )
            self.finished.emit(xls_path, temp_dir)
        except Exception as exc:
            self.error.emit(str(exc))


class _PreviewWorker(QObject):
    finished = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, xls_path: str):
        super().__init__()
        self.xls_path = xls_path

    def run(self) -> None:
        try:
            excel_generator.open_label_live_preview(self.xls_path)
            self.finished.emit()
        except Exception as exc:
            self.error.emit(str(exc))


class _PrintWorker(QObject):
    finished = pyqtSignal(str)  # used printer
    error = pyqtSignal(str)

    def __init__(self, xls_path: str, printer_name: str, margins: dict):
        super().__init__()
        self.xls_path = xls_path
        self.printer_name = printer_name
        self.margins = margins

    def run(self) -> None:
        try:
            used = excel_generator.print_label_file(
                self.xls_path,
                printer_name=self.printer_name,
                margins=self.margins,
            )
            self.finished.emit(used)
        except Exception as exc:
            self.error.emit(str(exc))


class ProductLabelsTableDialog(QDialog):
    """Окно с таблицей этикеток продукта (при двойном клике)."""

    def __init__(
        self,
        parent,
        routes: list[dict],
        file_type: str,
        products_ref: list,
        departments_ref: list,
        product_name: str,
        dept_key: str | None,
    ):
        super().__init__(parent)
        self.setWindowTitle(f"Этикетки: {product_name}")
        self.setMinimumSize(560, 360)

        rows = excel_generator.labels_preview_rows(
            routes,
            file_type,
            products_ref,
            departments_ref,
            only_product=product_name,
            only_dept_key=dept_key,
        )

        lay = QVBoxLayout(self)
        lay.setContentsMargins(14, 12, 14, 12)
        table = QTableWidget()
        table.setColumnCount(5)
        table.setHorizontalHeaderLabels(
            ["№ маршрута", "Адрес", "Продукт", "Отдел", "Кол-во"]
        )
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        table.setRowCount(len(rows))
        for r, (route_num, address, prod_name, dept_name, qty) in enumerate(rows):
            table.setItem(r, 0, QTableWidgetItem(route_num))
            table.setItem(r, 1, QTableWidgetItem(address))
            table.setItem(r, 2, QTableWidgetItem(prod_name))
            table.setItem(r, 3, QTableWidgetItem(dept_name))
            table.setItem(r, 4, QTableWidgetItem(qty))
        lay.addWidget(table)
        btn = QPushButton("Закрыть")
        btn.setObjectName("btnSecondary")
        btn.clicked.connect(self.accept)
        lay.addWidget(btn)


class LabelsPrintPreviewDialog(QDialog):
    """Live-preview и печать этикеток одного продукта."""

    def __init__(
        self,
        parent,
        routes: list[dict],
        file_type: str,
        products_ref: list,
        departments_ref: list,
        product_name: str,
        dept_key: str | None,
    ):
        super().__init__(parent)
        self.setWindowTitle(f"Предпросмотр и печать: {product_name}")
        self.setMinimumSize(720, 460)

        self._routes = routes
        self._file_type = file_type
        self._products_ref = products_ref
        self._departments_ref = departments_ref
        self._product_name = product_name
        self._dept_key = dept_key
        self._temp_dir: str | None = None
        self._xls_path: str | None = None
        self._preview_running = False
        self._threads: list[QThread] = []

        self._build_ui()
        self._fill_preview_table()
        self._prepare_temp_file()

    def _build_ui(self) -> None:
        lay = QVBoxLayout(self)
        lay.setContentsMargins(18, 14, 18, 14)
        lay.setSpacing(10)

        self.lbl_status = QLabel("Подготовка предпросмотра...")
        self.lbl_status.setObjectName("hintLabel")
        lay.addWidget(self.lbl_status)

        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(
            ["№ маршрута", "Адрес", "Продукт", "Отдел", "Кол-во"]
        )
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        lay.addWidget(self.table, 1)

        printer_row = QHBoxLayout()
        printer_row.addWidget(QLabel("Принтер:"))
        self.combo_printer = QComboBox()
        self.combo_printer.setMinimumWidth(320)
        printer_row.addWidget(self.combo_printer, 1)
        self.btn_choose_printer = QPushButton("Выбрать другой...")
        self.btn_choose_printer.setObjectName("btnSecondary")
        self.btn_choose_printer.clicked.connect(self._on_choose_printer)
        printer_row.addWidget(self.btn_choose_printer)
        lay.addLayout(printer_row)

        btn_row = QHBoxLayout()
        self.btn_open_preview = QPushButton("Открыть live-preview Excel")
        self.btn_open_preview.setObjectName("btnSecondary")
        self.btn_open_preview.clicked.connect(self._on_open_preview)
        self.btn_open_preview.setEnabled(False)
        btn_row.addWidget(self.btn_open_preview)

        self.btn_print = QPushButton(f"Печать этикеток: {self._product_name}")
        self.btn_print.setObjectName("btnPrimary")
        self.btn_print.clicked.connect(self._on_print)
        self.btn_print.setEnabled(False)
        btn_row.addWidget(self.btn_print)
        btn_row.addStretch()

        btn_close = QPushButton("Закрыть")
        btn_close.setObjectName("btnSecondary")
        btn_close.clicked.connect(self.accept)
        btn_row.addWidget(btn_close)
        lay.addLayout(btn_row)

        self._reload_printers()

    def _fill_preview_table(self) -> None:
        rows = excel_generator.labels_preview_rows(
            self._routes,
            self._file_type,
            self._products_ref,
            self._departments_ref,
            only_product=self._product_name,
            only_dept_key=self._dept_key,
        )
        self.table.setRowCount(len(rows))
        for r, (route_num, address, prod_name, dept_name, qty) in enumerate(rows):
            self.table.setItem(r, 0, QTableWidgetItem(route_num))
            self.table.setItem(r, 1, QTableWidgetItem(address))
            self.table.setItem(r, 2, QTableWidgetItem(prod_name))
            self.table.setItem(r, 3, QTableWidgetItem(dept_name))
            self.table.setItem(r, 4, QTableWidgetItem(qty))

    def _reload_printers(self) -> None:
        self.combo_printer.clear()
        printers = excel_generator.get_excel_printers()
        for p in printers:
            self.combo_printer.addItem(p, p)
        remembered = str(data_store.get_setting("labelsLastPrinter") or "").strip()
        if remembered and self.combo_printer.findData(remembered) < 0:
            self.combo_printer.addItem(remembered, remembered)
        if remembered:
            idx = self.combo_printer.findData(remembered)
            if idx >= 0:
                self.combo_printer.setCurrentIndex(idx)

    def _prepare_temp_file(self) -> None:
        self.btn_open_preview.setEnabled(False)
        self.btn_print.setEnabled(False)
        thread = QThread(self)
        worker = _PrepareWorker(
            routes=self._routes,
            file_type=self._file_type,
            products_ref=self._products_ref,
            departments_ref=self._departments_ref,
            product_name=self._product_name,
            dept_key=self._dept_key,
        )
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.finished.connect(self._on_prepare_done)
        worker.error.connect(self._on_prepare_error)
        worker.finished.connect(thread.quit)
        worker.error.connect(thread.quit)
        thread.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        self._threads.append(thread)
        thread.start()

    def _on_prepare_done(self, xls_path: str, temp_dir: str) -> None:
        self._xls_path = xls_path
        self._temp_dir = temp_dir
        self.lbl_status.setText("Файл для preview/печати готов.")
        self.btn_open_preview.setEnabled(True)
        self.btn_print.setEnabled(True)

    def _on_prepare_error(self, msg: str) -> None:
        self.lbl_status.setText("Ошибка подготовки предпросмотра.")
        QMessageBox.critical(self, "Ошибка", msg)

    def _on_choose_printer(self) -> None:
        items = [self.combo_printer.itemData(i) for i in range(self.combo_printer.count())]
        items = [str(x) for x in items if str(x).strip()]
        if items:
            current = str(self.combo_printer.currentData() or "")
            item, ok = QInputDialog.getItem(
                self,
                "Выбор принтера",
                "Принтер:",
                items,
                max(0, items.index(current)) if current in items else 0,
                editable=False,
            )
            if ok and item:
                idx = self.combo_printer.findData(item)
                if idx >= 0:
                    self.combo_printer.setCurrentIndex(idx)
                data_store.set_setting("labelsLastPrinter", item)
            return
        text, ok = QInputDialog.getText(self, "Выбор принтера", "Введите имя принтера:")
        if ok and (text or "").strip():
            val = text.strip()
            if self.combo_printer.findData(val) < 0:
                self.combo_printer.addItem(val, val)
            self.combo_printer.setCurrentIndex(self.combo_printer.findData(val))
            data_store.set_setting("labelsLastPrinter", val)

    def _on_open_preview(self) -> None:
        if not self._xls_path or self._preview_running:
            return
        self._preview_running = True
        self.btn_open_preview.setEnabled(False)
        self.lbl_status.setText("Открыт live-preview Excel. Закройте книгу, чтобы вернуться.")
        thread = QThread(self)
        worker = _PreviewWorker(self._xls_path)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.finished.connect(self._on_preview_done)
        worker.error.connect(self._on_preview_error)
        worker.finished.connect(thread.quit)
        worker.error.connect(thread.quit)
        thread.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        self._threads.append(thread)
        thread.start()

    def _on_preview_done(self) -> None:
        self._preview_running = False
        self.btn_open_preview.setEnabled(bool(self._xls_path))
        self.lbl_status.setText("Live-preview закрыт.")

    def _on_preview_error(self, msg: str) -> None:
        self._preview_running = False
        self.btn_open_preview.setEnabled(bool(self._xls_path))
        self.lbl_status.setText("Ошибка live-preview.")
        QMessageBox.warning(self, "Не удалось открыть live-preview", msg)

    def _on_print(self) -> None:
        if not self._xls_path:
            return
        printer = str(self.combo_printer.currentData() or "").strip()
        margins = dict(data_store.get_setting("labelsPrintMargins") or {})
        if not margins:
            margins = {
                "top_cm": 2.0,
                "right_cm": 2.0,
                "bottom_cm": 0.0,
                "left_cm": 0.0,
            }
        self.btn_print.setEnabled(False)
        self.lbl_status.setText("Отправка в печать...")
        thread = QThread(self)
        worker = _PrintWorker(self._xls_path, printer, margins)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.finished.connect(self._on_print_done)
        worker.error.connect(self._on_print_error)
        worker.finished.connect(thread.quit)
        worker.error.connect(thread.quit)
        thread.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        self._threads.append(thread)
        thread.start()

    def _on_print_done(self, used_printer: str) -> None:
        self.btn_print.setEnabled(True)
        if used_printer:
            data_store.set_setting("labelsLastPrinter", used_printer)
        self.lbl_status.setText(f"Печать отправлена. Принтер: {used_printer or 'по умолчанию'}")
        QMessageBox.information(
            self,
            "Печать",
            f"Этикетки отправлены на печать.\nПринтер: {used_printer or 'по умолчанию'}\n"
            "Отступы: верх 2 см, справа 2 см, снизу 0 см, слева 0 см.",
        )

    def _on_print_error(self, msg: str) -> None:
        self.btn_print.setEnabled(True)
        self.lbl_status.setText("Ошибка печати.")
        QMessageBox.critical(self, "Ошибка печати", msg)

    def closeEvent(self, event) -> None:
        try:
            auto_cleanup = bool(data_store.get_setting("labelsTempAutoCleanup"))
            if auto_cleanup and self._temp_dir and os.path.isdir(self._temp_dir):
                shutil.rmtree(self._temp_dir, ignore_errors=True)
        finally:
            super().closeEvent(event)

