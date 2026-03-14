"""
dashboard_page.py — Главная страница приложения (дашборд).

Рекомендованный вариант A: сетка из 6 карточек быстрых действий.
"""
from __future__ import annotations

import sys
from pathlib import Path
if __name__ == "__main__":
    _root = Path(__file__).resolve().parents[2]
    sys.path.insert(0, str(_root))

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFrame, QGridLayout, QScrollArea, QSizePolicy,
)
from PyQt6.QtCore import Qt, pyqtSignal

from core import data_store


class DashboardPage(QWidget):
    """Главная страница с карточками быстрых действий."""

    open_history = pyqtSignal()
    go_process_files = pyqtSignal()
    go_last_main = pyqtSignal()
    go_last_increase = pyqtSignal()
    go_labels = pyqtSignal()
    go_clear = pyqtSignal()

    def __init__(self, app_state: dict):
        super().__init__()
        self.app_state = app_state
        self._build_ui()

    def _build_ui(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        content = QWidget()
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.addWidget(scroll)
        inner = QVBoxLayout(content)
        inner.setContentsMargins(24, 20, 24, 20)
        inner.setSpacing(16)
        scroll.setWidget(content)

        title = QLabel("Маршруты, Сборка")
        title.setObjectName("sectionTitle")
        inner.addWidget(title)

        hint = QLabel(
            "Быстрый доступ к основным действиям. Обработка файлов и этикетки — через вкладки ленты или карточки ниже."
        )
        hint.setObjectName("stepLabel")
        hint.setWordWrap(True)
        inner.addWidget(hint)

        grid = QGridLayout()
        grid.setSpacing(10)

        self._card_history = self._make_card(
            "📋", "История",
            "Открыть историю маршрутов (Основные и Увеличение)",
            "btnPrimary", self.open_history.emit, "history"
        )
        grid.addWidget(self._card_history, 0, 0)

        self._card_process = self._make_card(
            "📂", "Обработать файлы",
            "Загрузить XLS-файлы и создать маршруты",
            "btnSecondary", self.go_process_files.emit, "process"
        )
        grid.addWidget(self._card_process, 0, 1)

        self._card_last_main = self._make_card(
            "📄", "Последние (основной)",
            "Открыть последние сохранённые маршруты основного типа",
            "btnSecondary", self.go_last_main.emit, "last_main"
        )
        grid.addWidget(self._card_last_main, 0, 2)

        self._card_last_inc = self._make_card(
            "📄", "Последние (довоз)",
            "Открыть последние сохранённые маршруты довоза",
            "btnSecondary", self.go_last_increase.emit, "last_inc"
        )
        grid.addWidget(self._card_last_inc, 1, 0)

        self._card_labels = self._make_card(
            "🏷️", "Этикетки",
            "Создать этикетки по шаблонам продуктов",
            "btnSecondary", self.go_labels.emit, "labels"
        )
        grid.addWidget(self._card_labels, 1, 1)

        self._card_clear = self._make_card(
            "🗑️", "Очистить",
            "Удалить сохранённые маршруты из памяти",
            "btnDanger", self.go_clear.emit, "clear"
        )
        grid.addWidget(self._card_clear, 1, 2)

        inner.addLayout(grid)
        inner.addStretch()

    def _make_card(self, icon: str, title: str, desc: str,
                   btn_style: str, on_click, card_key: str) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        card.setProperty("cardKey", card_key)
        card.setCursor(Qt.CursorShape.PointingHandCursor)
        card.setMinimumWidth(180)
        card_lay = QVBoxLayout(card)
        card_lay.setContentsMargins(16, 12, 16, 12)
        card_lay.setSpacing(8)

        lbl_icon = QLabel(icon)
        lbl_icon.setObjectName("dropZoneIcon")
        lbl_icon.setAlignment(Qt.AlignmentFlag.AlignLeft)
        lbl_icon.setStyleSheet("font-size: 22px;")
        card_lay.addWidget(lbl_icon)

        lbl_title = QLabel(title)
        lbl_title.setObjectName("cardTitle")
        lbl_title.setWordWrap(True)
        card_lay.addWidget(lbl_title)

        lbl_desc = QLabel(desc)
        lbl_desc.setObjectName("stepLabel")
        lbl_desc.setWordWrap(True)
        lbl_desc.setStyleSheet("font-size: 10px;")
        card_lay.addWidget(lbl_desc)

        card_lay.addStretch()

        btn = QPushButton(title)
        btn.setObjectName(btn_style)
        btn.setFixedHeight(32)
        btn.setMinimumWidth(160)
        btn.setSizePolicy(QSizePolicy.Policy.MinimumExpanding, QSizePolicy.Policy.Fixed)
        btn.clicked.connect(on_click)
        btn.setProperty("cardKey", card_key)
        card_lay.addWidget(btn)

        return card

    def refresh(self):
        """Обновляет подсказки и доступность карточек."""
        main_hist = data_store.get_routes_history("main")
        inc_hist = data_store.get_routes_history("increase")
        total = len(main_hist) + len(inc_hist)
        main_data = data_store.get_last_routes("main")
        inc_data = data_store.get_last_routes("increase")
        has_main = main_data is not None and bool(main_data.get("routes") or main_data.get("filteredRoutes"))
        has_inc = inc_data is not None and bool(inc_data.get("routes") or inc_data.get("filteredRoutes"))

        # История
        if total > 0:
            tip = f"Основные: {len(main_hist)}, Увеличение: {len(inc_hist)}. Выберите сохранение из списка."
            self._set_card_enabled(self._card_history, True, tip)
        else:
            self._set_card_enabled(self._card_history, False, "История пуста. Сначала обработайте файлы.")

        # Последние (основной)
        self._set_card_enabled(
            self._card_last_main,
            has_main,
            "Открыть последние маршруты основного типа" if has_main else "Нет сохранённых маршрутов. Сначала обработайте файлы."
        )

        # Последние (довоз)
        self._set_card_enabled(
            self._card_last_inc,
            has_inc,
            "Открыть последние маршруты довоза" if has_inc else "Нет сохранённых маршрутов. Сначала обработайте файлы."
        )

    def _set_card_enabled(self, card: QFrame, enabled: bool, tooltip: str):
        card.setToolTip(tooltip)
        for w in card.findChildren(QPushButton):
            w.setEnabled(enabled)
            w.setToolTip(tooltip)
            break
