from __future__ import annotations

import logging
import os
import shutil

from PyQt6.QtCore import QObject, QThread, pyqtSignal, Qt
from PyQt6.QtGui import QImage, QPixmap
from PyQt6.QtWidgets import (
    QApplication,
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
    QScrollArea,
    QFrame,
    QSplitter,
)

from core import data_store, excel_generator

log = logging.getLogger("labels_preview")


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


class _ExportPdfWorker(QObject):
    """Экспорт XLS в PDF для точного предпросмотра."""
    finished = pyqtSignal(str)  # pdf_path
    error = pyqtSignal(str)

    def __init__(self, xls_path: str, output_pdf_path: str, margins: dict):
        super().__init__()
        self.xls_path = xls_path
        self.output_pdf_path = output_pdf_path
        self.margins = margins

    def run(self) -> None:
        try:
            excel_generator.export_label_to_pdf(
                self.xls_path,
                self.output_pdf_path,
                margins=self.margins,
            )
            self.finished.emit(self.output_pdf_path)
        except Exception as exc:
            self.error.emit(str(exc))


def _pdf_first_page_to_pixmap(pdf_path: str, dpi: int = 150) -> QPixmap | None:
    """Рендерит первую страницу PDF в QPixmap. Возвращает None при ошибке."""
    if not pdf_path or not os.path.isfile(pdf_path):
        log.warning("PDF файл не найден: %s", pdf_path)
        return None
    try:
        import fitz
        import tempfile
        doc = fitz.open(pdf_path)
        try:
            if doc.page_count == 0:
                log.warning("PDF пустой (0 страниц)")
                return None
            page = doc[0]
            pix = page.get_pixmap(dpi=dpi, alpha=False)
            if pix.width <= 0 or pix.height <= 0:
                log.warning("Pixmap пустой: %dx%d", pix.width, pix.height)
                return None
            log.debug("PDF pixmap: %dx%d", pix.width, pix.height)
            # Способ 1: save в temp PNG (наиболее надёжно для QImage)
            fd, png_path = tempfile.mkstemp(suffix=".png")
            try:
                os.close(fd)
                pix.save(png_path)
                img = QImage(png_path)
                if not img.isNull():
                    result = QPixmap.fromImage(img)
                    log.debug("Предпросмотр: загружен из temp PNG")
                    return result
            except Exception as e1:
                log.debug("pix.save(png) не сработал: %s", e1)
            finally:
                if os.path.isfile(png_path):
                    try:
                        os.unlink(png_path)
                    except Exception:
                        pass
            # Способ 2: tobytes PNG
            for fn in [lambda: pix.tobytes("png"), lambda: pix.tobytes(output="png")]:
                try:
                    png_bytes = fn()
                    img = QImage()
                    if img.loadFromData(png_bytes):
                        log.debug("Предпросмотр: загружен из tobytes PNG")
                        return QPixmap.fromImage(img)
                except Exception as e2:
                    log.debug("tobytes PNG: %s", e2)
            # Способ 3: Pillow
            try:
                png_bytes = pix.pil_tobytes(format="PNG")
                img = QImage()
                if img.loadFromData(png_bytes):
                    log.debug("Предпросмотр: загружен через pil_tobytes")
                    return QPixmap.fromImage(img)
            except Exception as e3:
                log.debug("pil_tobytes: %s", e3)
            # Способ 4: raw RGB
            samples = bytes(pix.samples) if not isinstance(pix.samples, bytes) else pix.samples
            stride = pix.stride if pix.stride > 0 else pix.width * 3
            img = QImage(samples, pix.width, pix.height, stride, QImage.Format.Format_RGB888)
            if not img.isNull():
                log.debug("Предпросмотр: raw RGB")
                return QPixmap.fromImage(img.copy())
            return None
        finally:
            doc.close()
    except ImportError as e:
        log.warning("PyMuPDF не установлен — точный предпросмотр недоступен: %s", e)
        return None
    except Exception as exc:
        log.warning("Ошибка рендера PDF %s: %s", pdf_path, exc, exc_info=True)
        return None


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
        self.setMinimumSize(800, 620)
        self.resize(900, 720)

        self._routes = routes
        self._file_type = file_type
        self._products_ref = products_ref
        self._departments_ref = departments_ref
        self._product_name = product_name
        self._dept_key = dept_key
        self._temp_dir: str | None = None
        self._xls_path: str | None = None
        self._pdf_path: str | None = None
        self._preview_running = False
        self._threads: list[QThread] = []
        self._closed = False

        self._build_ui()
        self._fill_preview_table()
        self._prepare_temp_file()

    def closeEvent(self, event) -> None:
        self._closed = True
        super().closeEvent(event)

    def _build_ui(self) -> None:
        lay = QVBoxLayout(self)
        lay.setContentsMargins(18, 14, 18, 14)
        lay.setSpacing(10)

        self.lbl_status = QLabel("Подготовка предпросмотра...")
        self.lbl_status.setObjectName("hintLabel")
        lay.addWidget(self.lbl_status)

        splitter = QSplitter(Qt.Orientation.Vertical)

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
        self.table.setMinimumHeight(120)
        splitter.addWidget(self.table)

        preview_frame = QFrame()
        preview_frame.setObjectName("card")
        preview_lay = QVBoxLayout(preview_frame)
        preview_lay.setContentsMargins(8, 8, 8, 8)
        preview_lay.addWidget(QLabel("Точный предпросмотр этикетки (как при печати):"))
        self.preview_scroll = QScrollArea()
        self.preview_scroll.setWidgetResizable(True)
        self.preview_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.preview_scroll.setMinimumHeight(200)
        self.preview_lbl = QLabel()
        self.preview_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_lbl.setMinimumSize(400, 200)
        self.preview_lbl.setScaledContents(False)
        self.preview_lbl.setText("Экспорт в PDF...")
        self.preview_lbl.setObjectName("hintLabel")
        self.preview_scroll.setWidget(self.preview_lbl)
        preview_lay.addWidget(self.preview_scroll, 1)
        splitter.addWidget(preview_frame)

        splitter.setSizes([280, 400])
        lay.addWidget(splitter, 1)

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
        if self._closed:
            return
        self._xls_path = xls_path
        self._temp_dir = temp_dir
        self.lbl_status.setText("Файл готов. Экспорт в PDF для предпросмотра...")
        self.btn_open_preview.setEnabled(True)
        self.btn_print.setEnabled(True)
        self._start_export_pdf()

    def _start_export_pdf(self) -> None:
        if not self._xls_path or not self._temp_dir:
            return
        pdf_path = os.path.join(self._temp_dir, "preview.pdf")
        margins = dict(data_store.get_setting("labelsPrintMargins") or {})
        if not margins:
            margins = {"top_cm": 2.0, "right_cm": 2.0, "bottom_cm": 0.0, "left_cm": 0.0}
        thread = QThread(self)
        worker = _ExportPdfWorker(self._xls_path, pdf_path, margins)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.finished.connect(self._on_export_pdf_done)
        worker.error.connect(self._on_export_pdf_error)
        worker.finished.connect(thread.quit)
        worker.error.connect(thread.quit)
        thread.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        self._threads.append(thread)
        thread.start()

    def _on_export_pdf_done(self, pdf_path: str) -> None:
        if self._closed:
            return
        self._pdf_path = pdf_path
        self.lbl_status.setText("Файл для preview/печати готов.")
        log.info("PDF экспорт готов: %s", pdf_path)
        try:
            pix = _pdf_first_page_to_pixmap(pdf_path, dpi=150)
        except Exception as exc:
            log.warning("Ошибка рендера PDF: %s", exc, exc_info=True)
            pix = None
        if pix and not pix.isNull():
            if pix.width() > 750:
                pix = pix.scaledToWidth(750, Qt.TransformationMode.SmoothTransformation)
            self.preview_lbl.setText("")
            self.preview_lbl.setPixmap(pix)
            self.preview_lbl.adjustSize()
            self.preview_lbl.update()
            log.info("Предпросмотр отображён: %dx%d", pix.width(), pix.height())
        else:
            err_hint = "Проверьте: 1) PyMuPDF установлен (pip install pymupdf), 2) PDF создан"
            if pdf_path and os.path.isfile(pdf_path):
                err_hint += f"\nPDF создан: {os.path.basename(pdf_path)} — ошибка рендера."
            else:
                err_hint += " — PDF не создан."
            self.preview_lbl.setText(err_hint)
            log.warning("Не удалось отобразить предпросмотр PDF: %s", pdf_path)

    def _on_export_pdf_error(self, msg: str) -> None:
        if self._closed:
            return
        self.lbl_status.setText("Файл для preview/печати готов.")
        hint = f"Предпросмотр недоступен: {msg}"
        hint += "\n\nДля экспорта в PDF требуется Microsoft Excel."
        self.preview_lbl.setText(hint)
        log.warning("Экспорт этикеток в PDF: %s", msg)

    def _on_prepare_error(self, msg: str) -> None:
        if self._closed:
            return
        self.lbl_status.setText("Ошибка подготовки предпросмотра.")
        self.preview_lbl.setText(f"Не удалось подготовить файл этикеток:\n\n{msg}")
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

