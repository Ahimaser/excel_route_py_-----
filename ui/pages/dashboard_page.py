"""
dashboard_page.py — Главная страница приложения (дашборд).
"""
from __future__ import annotations

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFrame, QMessageBox, QGridLayout,
)
from PyQt6.QtCore import Qt, pyqtSignal

from core import data_store


class DashboardPage(QWidget):
    """Главная страница с карточками действий."""

    go_process_files  = pyqtSignal()
    open_last_main    = pyqtSignal()
    open_last_increase = pyqtSignal()
    go_labels         = pyqtSignal()
    clear_last_routes = pyqtSignal()

    def __init__(self, app_state: dict):
        super().__init__()
        self.app_state = app_state
        self._build_ui()

    # ─────────────────────────── UI ───────────────────────────────────

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(48, 40, 48, 40)
        lay.setSpacing(32)

        # Заголовок
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

        # Сетка карточек 2×2
        grid = QGridLayout()
        grid.setSpacing(16)

        self._card_process = self._make_card(
            "📂", "Обработать файлы",
            "Загрузить XLS-файлы маршрутов (ШК и/или СД)",
            "btnPrimary", self.go_process_files.emit
        )
        grid.addWidget(self._card_process, 0, 0)

        self._card_last_main = self._make_card(
            "📋", "Последние (основной)",
            "Открыть последние сохранённые маршруты",
            "btnSecondary", self.open_last_main.emit
        )
        grid.addWidget(self._card_last_main, 0, 1)

        self._card_last_inc = self._make_card(
            "🔄", "Последние (довоз)",
            "Открыть последние сохранённые маршруты (увеличение)",
            "btnSecondary", self.open_last_increase.emit
        )
        grid.addWidget(self._card_last_inc, 1, 0)

        self._card_labels = self._make_card(
            "🏷", "Этикетки",
            "Создать этикетки XLS по шаблонам продуктов",
            "btnSecondary", self.go_labels.emit
        )
        grid.addWidget(self._card_labels, 1, 1)

        lay.addLayout(grid)

        # Кнопка очистки внизу
        clear_row = QHBoxLayout()
        btn_clear = QPushButton("Очистить последние данные")
        btn_clear.setObjectName("btnDanger")
        btn_clear.setFixedHeight(40)
        btn_clear.setToolTip("Удалить сохранённые маршруты (если загружен неправильный файл)")
        btn_clear.clicked.connect(self._on_clear_last)
        clear_row.addStretch()
        clear_row.addWidget(btn_clear)
        lay.addLayout(clear_row)

        lay.addStretch()

    def _make_card(self, icon: str, title: str, desc: str,
                   btn_style: str, on_click) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        card.setCursor(Qt.CursorShape.PointingHandCursor)
        card_lay = QVBoxLayout(card)
        card_lay.setContentsMargins(24, 20, 24, 20)
        card_lay.setSpacing(12)

        lbl_icon = QLabel(icon)
        lbl_icon.setObjectName("dropZoneIcon")
        lbl_icon.setAlignment(Qt.AlignmentFlag.AlignLeft)
        lbl_icon.setStyleSheet("font-size: 28px;")
        card_lay.addWidget(lbl_icon)

        lbl_title = QLabel(title)
        lbl_title.setObjectName("cardTitle")
        card_lay.addWidget(lbl_title)

        lbl_desc = QLabel(desc)
        lbl_desc.setObjectName("stepLabel")
        lbl_desc.setWordWrap(True)
        lbl_desc.setStyleSheet("font-size: 12px;")
        card_lay.addWidget(lbl_desc)

        card_lay.addStretch()

        btn = QPushButton(title)
        btn.setObjectName(btn_style)
        btn.setFixedHeight(40)
        btn.clicked.connect(on_click)
        card_lay.addWidget(btn)

        return card

    # ─────────────────────────── Обновление ───────────────────────────

    def refresh(self):
        """Обновляет подсказки карточек «Последние» по состоянию хранилища."""
        main_data = data_store.get_last_routes("main")
        inc_data  = data_store.get_last_routes("increase")

        if main_data:
            n = len(main_data.get("filteredRoutes") or main_data.get("routes") or [])
            ts = (main_data.get("timestamp") or "")[:10]
            tip = f"Маршрутов: {n}" + (f" | {ts}" if ts else "")
            self._card_last_main.setToolTip(tip)
        else:
            self._card_last_main.setToolTip("Нет сохранённых данных")

        if inc_data:
            n = len(inc_data.get("filteredRoutes") or inc_data.get("routes") or [])
            ts = (inc_data.get("timestamp") or "")[:10]
            tip = f"Маршрутов: {n}" + (f" | {ts}" if ts else "")
            self._card_last_inc.setToolTip(tip)
        else:
            self._card_last_inc.setToolTip("Нет сохранённых данных")

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
