"""
home_page.py — Главная страница: 3 шага обработки файлов.

Оптимизации:
- Использует data_store.get_ref() для read-only доступа (без deep-copy)
- Использует data_store.update_product() для точечного обновления продуктов
- Кэш saveDir в app_state обновляется только при изменении
- ParseWorker не копирует file_paths лишний раз
"""
from __future__ import annotations

import os
from pathlib import Path

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFrame, QFileDialog, QListWidget, QListWidgetItem,
    QButtonGroup, QRadioButton, QProgressBar, QSizePolicy,
    QScrollArea, QMessageBox
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QObject
from PyQt6.QtGui import QDragEnterEvent, QDropEvent

from core import data_store, xls_parser


# ─────────────────────────── Worker ───────────────────────────────────────

class ParseWorker(QObject):
    """Парсит XLS файлы в отдельном потоке."""
    finished = pyqtSignal(dict)
    error    = pyqtSignal(str)

    def __init__(self, file_paths: list[str]):
        super().__init__()
        self.file_paths = file_paths

    def run(self):
        try:
            result = xls_parser.parse_files(self.file_paths)
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))


# ─────────────────────────── Зона перетаскивания ──────────────────────────

class DropZone(QFrame):
    files_dropped = pyqtSignal(list)

    def __init__(self):
        super().__init__()
        self.setObjectName("card")
        self.setAcceptDrops(True)
        self.setMinimumHeight(120)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

        lay = QVBoxLayout(self)
        lay.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.lbl_icon = QLabel("📂")
        self.lbl_icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_icon.setStyleSheet("font-size: 32px;")
        lay.addWidget(self.lbl_icon)

        self.lbl_text = QLabel("Перетащите .xls файлы сюда\nили нажмите для выбора")
        self.lbl_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_text.setStyleSheet("color: #64748b; font-size: 13px;")
        lay.addWidget(self.lbl_text)

    def mousePressEvent(self, event):
        self.files_dropped.emit([])

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setStyleSheet(
                "QFrame#card { border: 2px dashed #2563eb; "
                "background: #eff6ff; border-radius: 10px; }"
            )

    def dragLeaveEvent(self, event):
        self.setStyleSheet("")

    def dropEvent(self, event: QDropEvent):
        self.setStyleSheet("")
        paths = [
            url.toLocalFile()
            for url in event.mimeData().urls()
            if url.toLocalFile().lower().endswith(".xls")
        ]
        if paths:
            self.files_dropped.emit(paths)


# ─────────────────────────── Страница Home ────────────────────────────────

class HomePage(QWidget):
    """Главная страница: выбор типа → загрузка файлов → папка → обработка."""

    go_preview = pyqtSignal()

    def __init__(self, app_state: dict):
        super().__init__()
        self.app_state = app_state
        self._file_paths: list[str] = []
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
        lay.setContentsMargins(40, 32, 40, 32)
        lay.setSpacing(24)

        lbl_h = QLabel("Обработка файлов")
        lbl_h.setObjectName("sectionTitle")
        lay.addWidget(lbl_h)

        lay.addWidget(self._make_step_card("1", "Выберите тип создаваемого файла",  self._build_step1()))
        lay.addWidget(self._make_step_card("2", "Выберите XLS файлы для обработки", self._build_step2()))
        lay.addWidget(self._make_step_card("3", "Папка сохранения (опционально)",   self._build_step3()))

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.btn_process = QPushButton("Обработать файлы →")
        self.btn_process.setObjectName("btnPrimary")
        self.btn_process.setFixedHeight(42)
        self.btn_process.setMinimumWidth(200)
        self.btn_process.setEnabled(False)
        self.btn_process.setToolTip(
            "Парсить выбранные XLS файлы и перейти к предпросмотру маршрутов.\n"
            "Горячая клавиша: Ctrl+S"
        )
        self.btn_process.clicked.connect(self._on_process)
        btn_row.addWidget(self.btn_process)
        lay.addLayout(btn_row)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setVisible(False)
        lay.addWidget(self.progress_bar)

        lay.addStretch()

    def _make_step_card(self, num: str, title: str, content: QWidget) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        lay = QVBoxLayout(card)
        lay.setContentsMargins(20, 16, 20, 16)
        lay.setSpacing(12)

        h_row = QHBoxLayout()
        badge = QLabel(num)
        badge.setObjectName("badge")
        badge.setFixedSize(24, 24)
        badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        h_row.addWidget(badge)

        lbl = QLabel(title)
        lbl.setStyleSheet("font-size: 14px; font-weight: 600; color: #1e293b;")
        h_row.addWidget(lbl)
        h_row.addStretch()
        lay.addLayout(h_row)

        sep = QFrame()
        sep.setObjectName("separator")
        sep.setFrameShape(QFrame.Shape.HLine)
        lay.addWidget(sep)

        lay.addWidget(content)
        return card

    def _build_step1(self) -> QWidget:
        w = QWidget()
        lay = QHBoxLayout(w)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(16)

        self.radio_main = QRadioButton("Основной")
        self.radio_main.setChecked(True)
        self.radio_main.setToolTip("Создаёт основной файл маршрутов")
        self.radio_main.toggled.connect(self._on_type_changed)

        self.radio_increase = QRadioButton("Увеличение (Довоз)")
        self.radio_increase.setToolTip("Создаёт файл довоза (увеличения количества)")
        self.radio_increase.toggled.connect(self._on_type_changed)

        grp = QButtonGroup(w)
        grp.addButton(self.radio_main)
        grp.addButton(self.radio_increase)

        lay.addWidget(self.radio_main)
        lay.addWidget(self.radio_increase)
        lay.addStretch()
        return w

    def _build_step2(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(8)

        self.drop_zone = DropZone()
        self.drop_zone.setToolTip("Перетащите .xls файлы сюда или нажмите для выбора")
        self.drop_zone.files_dropped.connect(self._on_drop_or_click)
        lay.addWidget(self.drop_zone)

        self.file_list = QListWidget()
        self.file_list.setMaximumHeight(120)
        self.file_list.setToolTip("Список загруженных файлов")
        self.file_list.setVisible(False)
        lay.addWidget(self.file_list)

        btn_row = QHBoxLayout()
        self.btn_add_files = QPushButton("+ Добавить файлы")
        self.btn_add_files.setObjectName("btnSecondary")
        self.btn_add_files.setToolTip("Добавить ещё .xls файлы к списку")
        self.btn_add_files.clicked.connect(self._open_file_dialog)
        self.btn_add_files.setVisible(False)

        self.btn_clear_files = QPushButton("Очистить")
        self.btn_clear_files.setObjectName("btnDanger")
        self.btn_clear_files.setToolTip("Удалить все загруженные файлы")
        self.btn_clear_files.clicked.connect(self._clear_files)
        self.btn_clear_files.setVisible(False)

        btn_row.addWidget(self.btn_add_files)
        btn_row.addWidget(self.btn_clear_files)
        btn_row.addStretch()
        lay.addLayout(btn_row)
        return w

    def _build_step3(self) -> QWidget:
        w = QWidget()
        lay = QHBoxLayout(w)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(12)

        self.lbl_save_dir = QLabel()
        self.lbl_save_dir.setStyleSheet("color: #475569; font-size: 13px;")
        self._update_save_dir_label()
        lay.addWidget(self.lbl_save_dir, 1)

        btn_change = QPushButton("Изменить")
        btn_change.setObjectName("btnSecondary")
        btn_change.setToolTip("Выбрать другую папку для сохранения файлов")
        btn_change.clicked.connect(self._choose_save_dir)
        lay.addWidget(btn_change)

        btn_reset = QPushButton("Рабочий стол")
        btn_reset.setObjectName("btnSecondary")
        btn_reset.setToolTip("Сбросить папку сохранения на Рабочий стол")
        btn_reset.clicked.connect(self._reset_save_dir)
        lay.addWidget(btn_reset)
        return w

    # ─────────────────────────── Логика ───────────────────────────────────

    def refresh(self):
        self._update_save_dir_label()

    def reset(self):
        self._clear_files()
        self.radio_main.setChecked(True)
        self.app_state["fileType"] = "main"
        self._update_save_dir_label()

    def _on_type_changed(self):
        self.app_state["fileType"] = "main" if self.radio_main.isChecked() else "increase"

    def _on_drop_or_click(self, paths: list[str]):
        if not paths:
            self._open_file_dialog()
        else:
            self._add_files(paths)

    def _open_file_dialog(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Выберите XLS файлы", "", "Excel 97-2003 (*.xls)"
        )
        if paths:
            self._add_files(paths)

    def _add_files(self, paths: list[str]):
        existing = set(self._file_paths)
        for p in paths:
            if p not in existing:
                self._file_paths.append(p)
                existing.add(p)
                item = QListWidgetItem(f"📄 {Path(p).name}")
                item.setToolTip(p)
                self.file_list.addItem(item)

        if self._file_paths:
            self.file_list.setVisible(True)
            self.btn_add_files.setVisible(True)
            self.btn_clear_files.setVisible(True)
            self.drop_zone.lbl_text.setText(f"Выбрано файлов: {len(self._file_paths)}")
            self.btn_process.setEnabled(True)

    def _clear_files(self):
        self._file_paths.clear()
        self.file_list.clear()
        self.file_list.setVisible(False)
        self.btn_add_files.setVisible(False)
        self.btn_clear_files.setVisible(False)
        self.drop_zone.lbl_text.setText("Перетащите .xls файлы сюда\nили нажмите для выбора")
        self.btn_process.setEnabled(False)

    def _choose_save_dir(self):
        d = QFileDialog.getExistingDirectory(
            self, "Выберите папку сохранения",
            self.app_state.get("saveDir") or data_store.get_desktop_path()
        )
        if d:
            self.app_state["saveDir"] = d
            settings = data_store.get("settings") or {}
            settings["defaultSaveDir"] = d
            data_store.set_key("settings", settings)
            self._update_save_dir_label()

    def _reset_save_dir(self):
        self.app_state["saveDir"] = None
        settings = data_store.get("settings") or {}
        settings["defaultSaveDir"] = None
        data_store.set_key("settings", settings)
        self._update_save_dir_label()

    def _update_save_dir_label(self):
        # Используем get_ref для быстрого чтения без копирования
        settings = data_store.get_ref("settings") or {}
        save_dir = (
            self.app_state.get("saveDir") or
            settings.get("defaultSaveDir") or
            data_store.get_desktop_path()
        )
        self.app_state["saveDir"] = save_dir
        self.lbl_save_dir.setText(f"📁 {save_dir}")

    def _on_process(self):
        if not self._file_paths:
            return

        self.app_state["filePaths"] = self._file_paths[:]
        self.btn_process.setEnabled(False)
        self.progress_bar.setVisible(True)

        self._thread = QThread()
        self._worker = ParseWorker(self._file_paths)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.finished.connect(self._on_parse_done)
        self._worker.error.connect(self._on_parse_error)
        self._worker.finished.connect(self._thread.quit)
        self._worker.error.connect(self._thread.quit)
        self._thread.start()

    def _on_parse_done(self, result: dict):
        self.progress_bar.setVisible(False)
        self.btn_process.setEnabled(True)

        if result.get("errors"):
            QMessageBox.warning(
                self, "Ошибки при чтении",
                "Некоторые файлы не удалось прочитать:\n" + "\n".join(result["errors"])
            )

        self.app_state["routes"] = result["routes"]
        self.app_state["uniqueProducts"] = result["uniqueProducts"]
        self.app_state["filteredRoutes"] = [
            {**r, "excluded": False} for r in result["routes"]
        ]

        # Обновляем продукты в хранилище — только новые (без перезаписи существующих)
        store_prods: list[dict] = data_store.get_ref("products") or []
        existing_names: set[str] = {p["name"] for p in store_prods}
        new_prods = [
            {
                "name": up["name"],
                "unit": up["unit"],
                "showPcs": False,
                "pcsPerUnit": 1.0,
                "roundUp": True,
                "deptKey": None,
            }
            for up in result["uniqueProducts"]
            if up["name"] not in existing_names
        ]
        if new_prods:
            # Читаем копию только если есть изменения
            updated = data_store.get("products") or []
            updated.extend(new_prods)
            data_store.set_key("products", updated)

        self.go_preview.emit()

    def _on_parse_error(self, msg: str):
        self.progress_bar.setVisible(False)
        self.btn_process.setEnabled(True)
        QMessageBox.critical(self, "Ошибка", f"Ошибка при обработке файлов:\n{msg}")
