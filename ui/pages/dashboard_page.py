"""
dashboard_page.py — Главная страница приложения (дашборд).
"""
from __future__ import annotations

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFrame, QMessageBox,
)
from PyQt6.QtCore import Qt, pyqtSignal


class DashboardPage(QWidget):
    """Главная страница с кнопками: Обработка файлов, Открыть последние, Этикетки."""

    go_process_files = pyqtSignal()
    open_last_main = pyqtSignal()
    open_last_increase = pyqtSignal()
    go_labels = pyqtSignal()
    clear_last_routes = pyqtSignal()

    def __init__(self, app_state: dict):
        super().__init__()
        self.app_state = app_state
        self._build_ui()

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(40, 32, 40, 32)
        lay.setSpacing(24)

        title = QLabel("Маршруты, Сборка")
        title.setObjectName("sectionTitle")
        lay.addWidget(title)

        hint = QLabel(
            "Обработайте XLS-файлы с маршрутами или откройте последние сохранённые данные. "
            "Этикетки создаются по шаблонам на странице «Этикетки»."
        )
        hint.setObjectName("stepLabel")
        hint.setWordWrap(True)
        lay.addWidget(hint)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(16)

        btn_open = QPushButton("Обработать файлы")
        btn_open.setObjectName("btnPrimary")
        btn_open.setFixedHeight(44)
        btn_open.setToolTip("Перейти к загрузке и обработке XLS-файлов (школы/сады)")
        btn_open.clicked.connect(self.go_process_files.emit)
        btn_row.addWidget(btn_open)

        btn_last_main = QPushButton("Последние (основной)")
        btn_last_main.setObjectName("btnSecondary")
        btn_last_main.setFixedHeight(44)
        btn_last_main.setToolTip("Открыть последние сохранённые маршруты (основной)")
        btn_last_main.clicked.connect(self.open_last_main.emit)
        btn_row.addWidget(btn_last_main)

        btn_last_inc = QPushButton("Последние (довоз)")
        btn_last_inc.setObjectName("btnSecondary")
        btn_last_inc.setFixedHeight(44)
        btn_last_inc.setToolTip("Открыть последние сохранённые маршруты (довоз)")
        btn_last_inc.clicked.connect(self.open_last_increase.emit)
        btn_row.addWidget(btn_last_inc)

        btn_labels = QPushButton("Этикетки")
        btn_labels.setObjectName("btnSecondary")
        btn_labels.setFixedHeight(44)
        btn_labels.setToolTip("Перейти к созданию этикеток XLS по шаблонам")
        btn_labels.clicked.connect(self.go_labels.emit)
        btn_row.addWidget(btn_labels)

        btn_clear = QPushButton("Очистить последние")
        btn_clear.setObjectName("btnDanger")
        btn_clear.setFixedHeight(44)
        btn_clear.setToolTip("Удалить сохранённые маршруты (если загружен неправильный файл)")
        btn_clear.clicked.connect(self._on_clear_last)
        btn_row.addWidget(btn_clear)

        btn_row.addStretch()
        lay.addLayout(btn_row)
        lay.addStretch()

    def _on_clear_last(self):
        reply = QMessageBox.question(
            self, "Очистить последние",
            "Удалить сохранённые маршруты (основной и довоз)?\n"
            "После этого кнопки «Последние» не будут открывать данные.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.clear_last_routes.emit()
