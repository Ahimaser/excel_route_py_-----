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
from datetime import date, timedelta
from pathlib import Path

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFrame, QFileDialog, QListWidget, QListWidgetItem,
    QButtonGroup, QRadioButton, QProgressBar, QSizePolicy,
    QScrollArea, QMessageBox, QComboBox, QDateEdit, QApplication,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QObject, QDate
from PyQt6.QtGui import QDragEnterEvent, QDropEvent

from core import data_store, xls_parser
from core.product_match import find_similar_canonicals


# ─────────────────────────── Worker ───────────────────────────────────────

class ParseWorker(QObject):
    """Парсит XLS файлы в отдельном потоке. file_categories: список «ШК»/«СД» по одному на файл."""
    finished = pyqtSignal(dict)
    error    = pyqtSignal(str)

    def __init__(
        self,
        file_paths: list[str],
        file_categories: list[str] | None = None,
    ):
        super().__init__()
        self.file_paths = file_paths
        self.file_categories = file_categories

    def run(self):
        try:
            result = xls_parser.parse_files(
                self.file_paths,
                file_categories=self.file_categories,
            )
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))


# ─────────────────────────── Зона перетаскивания ──────────────────────────

class DropZone(QFrame):
    files_dropped = pyqtSignal(list)

    def __init__(self):
        super().__init__()
        self.setObjectName("dropZoneCard")
        self.setAcceptDrops(True)
        self.setMinimumHeight(120)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

        lay = QVBoxLayout(self)
        lay.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.lbl_icon = QLabel("📂")
        self.lbl_icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_icon.setObjectName("dropZoneIcon")
        lay.addWidget(self.lbl_icon)

        self.lbl_text = QLabel("Перетащите .xls файлы сюда\nили нажмите для выбора")
        self.lbl_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_text.setObjectName("stepLabel")
        lay.addWidget(self.lbl_text)

    def mousePressEvent(self, event):
        self.files_dropped.emit([])

    def _set_drop_hover(self, hover: bool):
        self.setProperty("dropZoneHover", hover)
        self.style().unpolish(self)
        self.style().polish(self)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self._set_drop_hover(True)

    def dragLeaveEvent(self, event):
        self._set_drop_hover(False)

    def dropEvent(self, event: QDropEvent):
        self._set_drop_hover(False)
        paths = [
            url.toLocalFile()
            for url in event.mimeData().urls()
            if url.toLocalFile().lower().endswith(".xls")
        ]
        if paths:
            self.files_dropped.emit(paths)


# ─────────────────────────── Страница Home ────────────────────────────────

class HomePage(QWidget):
    """Главная страница: выбор типа → загрузка файлов (ШК/СД) → папка → обработка."""

    go_preview = pyqtSignal()
    go_dashboard = pyqtSignal()

    def __init__(self, app_state: dict):
        super().__init__()
        self.app_state = app_state
        self._file_paths_shk: list[str] = []
        self._file_paths_sd: list[str] = []
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
        lay.setContentsMargins(24, 16, 24, 16)
        lay.setSpacing(16)

        lbl_h = QLabel("Обработка файлов")
        lbl_h.setObjectName("sectionTitle")
        lay.addWidget(lbl_h)

        lay.addWidget(self._make_step_card("1", "Тип файла", self._build_step1()))
        lay.addWidget(self._make_step_card("2", "Файлы XLS (школы и/или сады)", self._build_step2()))
        lay.addWidget(self._make_step_card("3", "Папка сохранения и дата", self._build_step3()))

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.btn_process = QPushButton("Обработать файлы →")
        self.btn_process.setObjectName("btnPrimary")
        self.btn_process.setFixedHeight(34)
        self.btn_process.setMinimumWidth(180)
        self.btn_process.setEnabled(False)
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
        lay.setContentsMargins(16, 12, 16, 12)
        lay.setSpacing(10)

        h_row = QHBoxLayout()
        badge = QLabel(num)
        badge.setObjectName("badge")
        badge.setFixedSize(20, 20)
        badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        h_row.addWidget(badge)

        lbl = QLabel(title)
        lbl.setObjectName("cardTitle")
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
        lay.setSpacing(10)

        self.radio_main = QRadioButton("Основной")
        self.radio_main.setChecked(True)
        self.radio_main.toggled.connect(self._on_type_changed)

        self.radio_increase = QRadioButton("Увеличение (Довоз)")
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
        lay.setSpacing(10)

        lbl_shk = QLabel("Файлы для ШК (школы)")
        lbl_shk.setObjectName("subsectionLabel")
        lay.addWidget(lbl_shk)
        self.drop_zone_shk = DropZone()
        self.drop_zone_shk.files_dropped.connect(lambda paths: self._on_drop_or_click(paths, "shk"))
        lay.addWidget(self.drop_zone_shk)
        self.file_list_shk = QListWidget()
        self.file_list_shk.setMaximumHeight(100)
        self.file_list_shk.setVisible(False)
        lay.addWidget(self.file_list_shk)
        btn_shk = QHBoxLayout()
        self.btn_add_shk = QPushButton("+ Добавить файлы ШК")
        self.btn_add_shk.setObjectName("btnSecondary")
        self.btn_add_shk.setMinimumWidth(180)
        self.btn_add_shk.clicked.connect(lambda: self._open_file_dialog("shk"))
        self.btn_add_shk.setVisible(False)
        self.btn_clear_shk = QPushButton("Очистить ШК")
        self.btn_clear_shk.setObjectName("btnDanger")
        self.btn_clear_shk.clicked.connect(lambda: self._clear_files("shk"))
        self.btn_clear_shk.setVisible(False)
        btn_shk.addWidget(self.btn_add_shk)
        btn_shk.addWidget(self.btn_clear_shk)
        btn_shk.addStretch()
        lay.addLayout(btn_shk)

        lbl_sd = QLabel("Файлы для СД (сады)")
        lbl_sd.setObjectName("subsectionLabel")
        lay.addWidget(lbl_sd)
        self.drop_zone_sd = DropZone()
        self.drop_zone_sd.files_dropped.connect(lambda paths: self._on_drop_or_click(paths, "sd"))
        lay.addWidget(self.drop_zone_sd)
        self.file_list_sd = QListWidget()
        self.file_list_sd.setMaximumHeight(100)
        self.file_list_sd.setVisible(False)
        lay.addWidget(self.file_list_sd)
        btn_sd = QHBoxLayout()
        self.btn_add_sd = QPushButton("+ Добавить файлы СД")
        self.btn_add_sd.setObjectName("btnSecondary")
        self.btn_add_sd.setMinimumWidth(180)
        self.btn_add_sd.clicked.connect(lambda: self._open_file_dialog("sd"))
        self.btn_add_sd.setVisible(False)
        self.btn_clear_sd = QPushButton("Очистить СД")
        self.btn_clear_sd.setObjectName("btnDanger")
        self.btn_clear_sd.clicked.connect(lambda: self._clear_files("sd"))
        self.btn_clear_sd.setVisible(False)
        btn_sd.addWidget(self.btn_add_sd)
        btn_sd.addWidget(self.btn_clear_sd)
        btn_sd.addStretch()
        lay.addLayout(btn_sd)

        return w

    def _build_step3(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(8)

        row_dir = QHBoxLayout()
        self.lbl_save_dir = QLabel()
        self.lbl_save_dir.setObjectName("stepLabel")
        self.lbl_save_dir.setWordWrap(True)
        self._update_save_dir_label()
        row_dir.addWidget(self.lbl_save_dir, 1)

        btn_change = QPushButton("Изменить")
        btn_change.setObjectName("btnSecondary")
        btn_change.clicked.connect(self._choose_save_dir)
        row_dir.addWidget(btn_change)
        lay.addLayout(row_dir)

        row_date = QHBoxLayout()
        row_date.addWidget(QLabel("Дата в заголовках и названиях папок:"))
        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDisplayFormat("dd.MM.yyyy")
        tomorrow = date.today() + timedelta(days=1)
        self.date_edit.setDate(QDate(tomorrow.year, tomorrow.month, tomorrow.day))
        row_date.addWidget(self.date_edit)
        row_date.addStretch()
        lay.addLayout(row_date)

        return w

    # ─────────────────────────── Логика ───────────────────────────────────

    def refresh(self):
        self._update_save_dir_label()

    def reset(self):
        self._clear_files("shk")
        self._clear_files("sd")
        self.radio_main.setChecked(True)
        self.app_state["fileType"] = "main"
        self._update_save_dir_label()

    def _on_type_changed(self):
        self.app_state["fileType"] = "main" if self.radio_main.isChecked() else "increase"

    def _on_drop_or_click(self, paths: list[str], category: str):
        if not paths:
            self._open_file_dialog(category)
        else:
            self._add_files(paths, category)

    def _open_file_dialog(self, category: str):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Выберите XLS файлы", "", "Excel 97-2003 (*.xls)"
        )
        if paths:
            self._add_files(paths, category)

    def _add_files(self, paths: list[str], category: str):
        lst = self._file_paths_shk if category == "shk" else self._file_paths_sd
        list_w = self.file_list_shk if category == "shk" else self.file_list_sd
        drop = self.drop_zone_shk if category == "shk" else self.drop_zone_sd
        btn_add = self.btn_add_shk if category == "shk" else self.btn_add_sd
        btn_clear = self.btn_clear_shk if category == "shk" else self.btn_clear_sd
        existing = set(lst)
        for p in paths:
            if p not in existing:
                lst.append(p)
                existing.add(p)
                list_w.addItem(QListWidgetItem(f"📄 {Path(p).name}"))
        list_w.setVisible(True)
        btn_add.setVisible(True)
        btn_clear.setVisible(True)
        drop.lbl_text.setText(f"Выбрано файлов: {len(lst)}")
        self._update_process_button()

    def _clear_files(self, category: str):
        if category == "shk":
            self._file_paths_shk.clear()
            self.file_list_shk.clear()
            self.file_list_shk.setVisible(False)
            self.btn_add_shk.setVisible(False)
            self.btn_clear_shk.setVisible(False)
            self.drop_zone_shk.lbl_text.setText("Перетащите .xls файлы сюда\nили нажмите для выбора")
        else:
            self._file_paths_sd.clear()
            self.file_list_sd.clear()
            self.file_list_sd.setVisible(False)
            self.btn_add_sd.setVisible(False)
            self.btn_clear_sd.setVisible(False)
            self.drop_zone_sd.lbl_text.setText("Перетащите .xls файлы сюда\nили нажмите для выбора")
        self._update_process_button()

    def _update_process_button(self):
        total = len(self._file_paths_shk) + len(self._file_paths_sd)
        self.btn_process.setEnabled(total > 0)

    def _choose_save_dir(self):
        d = QFileDialog.getExistingDirectory(
            self, "Выберите папку сохранения",
            self.app_state.get("saveDir") or data_store.get_desktop_path()
        )
        if d:
            self.app_state["saveDir"] = d
            data_store.set_setting("defaultSaveDir", d)
            self._update_save_dir_label()

    def _reset_save_dir(self):
        self.app_state["saveDir"] = None
        data_store.set_setting("defaultSaveDir", None)
        self._update_save_dir_label()

    def _update_save_dir_label(self):
        save_dir = (
            self.app_state.get("saveDir") or
            data_store.get_setting("defaultSaveDir") or
            data_store.get_desktop_path()
        )
        self.app_state["saveDir"] = save_dir
        self.lbl_save_dir.setText(f"📁 {save_dir}")

    def _on_process(self):
        file_paths = self._file_paths_shk + self._file_paths_sd
        file_categories = ["ШК"] * len(self._file_paths_shk) + ["СД"] * len(self._file_paths_sd)
        if not file_paths:
            return

        # Проверка существования файлов перед парсингом
        existing_shk = [p for p in self._file_paths_shk if os.path.isfile(p)]
        existing_sd = [p for p in self._file_paths_sd if os.path.isfile(p)]
        missing = [p for p in file_paths if p not in existing_shk + existing_sd]
        if missing:
            msg = "Следующие файлы недоступны (удалены или перемещены) и будут пропущены:\n\n" + "\n".join(missing[:10])
            if len(missing) > 10:
                msg += f"\n\n... и ещё {len(missing) - 10}"
            QMessageBox.warning(self, "Файлы недоступны", msg)
        file_paths = existing_shk + existing_sd
        file_categories = ["ШК"] * len(existing_shk) + ["СД"] * len(existing_sd)
        if not file_paths:
            return

        self.app_state["filePaths"] = file_paths[:]
        self.btn_process.setEnabled(False)
        self.progress_bar.setVisible(True)
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)

        self._thread = QThread()
        self._worker = ParseWorker(file_paths, file_categories=file_categories)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.finished.connect(self._on_parse_done)
        self._worker.error.connect(self._on_parse_error)
        self._worker.finished.connect(self._thread.quit)
        self._worker.error.connect(self._thread.quit)
        self._thread.start()

    def _on_parse_done(self, result: dict):
        QApplication.restoreOverrideCursor()
        self.progress_bar.setVisible(False)
        self.btn_process.setEnabled(True)

        routes = result.get("routes") or []
        unique_products = result.get("uniqueProducts") or []
        if result.get("errors"):
            QMessageBox.warning(
                self, "Ошибки при чтении",
                "Некоторые файлы не удалось прочитать:\n" + "\n".join(result["errors"])
            )

        self.app_state["routes"] = routes
        self.app_state["uniqueProducts"] = unique_products
        self.app_state["filteredRoutes"] = [{**r, "excluded": False} for r in routes]
        self.app_state["institutionList"] = data_store.get_institution_list_from_routes(routes)
        first_cat = routes[0].get("routeCategory", "ШК") if routes else "ШК"
        qd = self.date_edit.date()
        self.app_state["routesDate"] = f"{qd.day():02d}.{qd.month():02d}.{qd.year()}"
        self.app_state["routeCategory"] = first_cat
        self.app_state["generalFileCreated"] = False
        self.app_state["deptFilesCreated"] = False

        if routes and hasattr(self.app_state.get("set_status"), "__call__"):
            n_prods = len(unique_products)
            self.app_state["set_status"](f"Загружено {len(routes)} маршрутов, {n_prods} продуктов")

        save_dir = self.app_state.get("saveDir") or data_store.get_setting("defaultSaveDir")
        data_store.save_last_routes(
            self.app_state.get("fileType", "main"),
            routes,
            unique_products,
            self.app_state["filteredRoutes"],
            route_category=first_cat,
            save_dir=save_dir,
        )

        store_prods = data_store.get_ref("products") or []
        existing_names = {p.get("name") for p in store_prods if p.get("name")}
        canonical_names = [p["name"] for p in store_prods if p.get("name") and p.get("deptKey")]

        new_items = []
        for up in unique_products:
            if not up.get("name") or up["name"] in existing_names:
                continue
            similar = find_similar_canonicals(up["name"], canonical_names)
            new_items.append({
                "name": up["name"],
                "unit": up.get("unit", ""),
                "similar": similar,
            })

        if new_items:
            try:
                from ui.pages.new_products_dialog import run_new_products_dialog
                decisions = run_new_products_dialog(self.window(), new_items, canonical_names)
            except Exception as e:
                import logging
                logging.getLogger("app").exception("Диалог новых продуктов: %s", e)
                decisions = [{"name": it["name"], "unit": it.get("unit", ""), "action": "new"} for it in new_items]
            updated = data_store.get("products") or []
            for d in decisions:
                if d["action"] == "alias":
                    data_store.set_alias(d["name"], d["canonical"])
                else:
                    updated.append({
                        "name": d["name"],
                        "unit": d.get("unit", ""),
                        "showPcs": False,
                        "pcsPerUnit": 1.0,
                        "roundUp": True,
                        "deptKey": None,
                    })
            if any(d["action"] == "new" for d in decisions):
                data_store.set_key("products", updated)

        self.go_preview.emit()

    def _on_parse_error(self, msg: str):
        QApplication.restoreOverrideCursor()
        self.progress_bar.setVisible(False)
        self.btn_process.setEnabled(True)
        QMessageBox.critical(self, "Ошибка", f"Ошибка при обработке файлов:\n{msg}")
