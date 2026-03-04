"""
labels_page.py — Страница создания этикеток XLS по шаблонам.
"""
from __future__ import annotations

import os
import logging
from datetime import date, timedelta

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFrame, QMessageBox, QScrollArea,
)
from PyQt6.QtCore import Qt, pyqtSignal

from core import data_store, excel_generator

log = logging.getLogger(__name__)


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
        btn_back.setToolTip("Вернуться на главную страницу")
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
            btn_settings.setToolTip("Отделы/подотделы: печать да/нет, шаблоны XLS по продуктам")
            btn_settings.clicked.connect(lambda: open_labels_settings_dialog(self))
            h.addWidget(btn_settings)
        except Exception:
            pass
        lay.addLayout(h)

        hint = QLabel(
            "Создайте этикетки из загруженных маршрутов по шаблонам XLS. "
            "В «Настройки этикеток» задаются печать по отделам и шаблон для каждого продукта. "
            "В последнюю строку шаблона подставляются: № маршрута, дом/строение, количество."
        )
        hint.setObjectName("stepLabel")
        hint.setWordWrap(True)
        lay.addWidget(hint)

        self.no_data_frame = QFrame()
        self.no_data_frame.setObjectName("card")
        no_lay = QVBoxLayout(self.no_data_frame)
        no_data_lbl = QLabel("Маршруты не загружены. Откройте последние или обработайте новые файлы.")
        no_data_lbl.setObjectName("stepLabel")
        no_lay.addWidget(no_data_lbl)
        btn_row = QHBoxLayout()
        btn_open = QPushButton("Открыть последние маршруты")
        btn_open.setObjectName("btnPrimary")
        btn_open.setToolTip("Загрузить последние сохранённые маршруты и перейти к предпросмотру")
        btn_open.clicked.connect(self.go_open_routes.emit)
        btn_process = QPushButton("Обработать файлы")
        btn_process.setObjectName("btnSecondary")
        btn_process.setToolTip("Перейти к загрузке и обработке XLS-файлов")
        btn_process.clicked.connect(self.go_process_files.emit)
        btn_row.addWidget(btn_open)
        btn_row.addWidget(btn_process)
        btn_row.addStretch()
        no_lay.addLayout(btn_row)
        lay.addWidget(self.no_data_frame)

        self.cards_frame = QFrame()
        cards_lay = QVBoxLayout(self.cards_frame)
        self.btn_xls = QPushButton("Создать XLS по шаблонам")
        self.btn_xls.setObjectName("btnPrimary")
        self.btn_xls.setFixedHeight(40)
        self.btn_xls.setToolTip("Этикетки сохраняются в папку «Этикетки на ДД.ММ.ГГГГ» (завтра).")
        self.btn_xls.clicked.connect(self._on_labels_from_templates)
        lbl_cards = QLabel("Этикетки из шаблонов (XLS)")
        lbl_cards.setObjectName("subsectionLabel")
        cards_lay.addWidget(lbl_cards)
        cards_lay.addWidget(self.btn_xls)
        lay.addWidget(self.cards_frame)
        lay.addStretch()

    def _has_routes(self) -> bool:
        routes = self.app_state.get("filteredRoutes", [])
        return bool([r for r in routes if not r.get("excluded")])

    def refresh(self):
        has = self._has_routes()
        self.no_data_frame.setVisible(not has)
        self.cards_frame.setVisible(has)

    def _on_labels_from_templates(self):
        routes = self.app_state.get("filteredRoutes", [])
        active = [r for r in routes if not r.get("excluded")]
        if not active:
            QMessageBox.warning(self, "Нет данных", "Нет маршрутов для этикеток.")
            return
        products_ref = data_store.get_ref("products") or []
        if not any(p.get("labelTemplatePath") for p in products_ref):
            QMessageBox.information(
                self, "Нет шаблонов",
                "Откройте «Настройки этикеток» и выберите шаблон XLS для продуктов."
            )
            return
        base_dir = self.app_state.get("saveDir") or data_store.get_desktop_path()
        tomorrow = date.today() + timedelta(days=1)
        folder_name = f"Этикетки на {tomorrow:%d.%m.%Y}"
        out_dir = os.path.join(base_dir, folder_name)
        os.makedirs(out_dir, exist_ok=True)
        file_type = self.app_state.get("fileType", "main")
        departments_ref = data_store.get_ref("departments") or []
        try:
            created = excel_generator.generate_labels_from_templates(
                routes, out_dir, file_type, products_ref, departments_ref
            )
            if created:
                QMessageBox.information(self, "Готово", f"Создано файлов: {len(created)}\n\n{out_dir}")
            else:
                QMessageBox.information(self, "Нет файлов", "Нет этикеток для создания.")
        except Exception as e:
            log.exception("labels")
            QMessageBox.critical(self, "Ошибка", str(e))
