"""
dashboard_page.py — Главная страница приложения (дашборд).
"""
from __future__ import annotations

import sys
from pathlib import Path
if __name__ == "__main__":
    # При запуске файла как скрипта добавляем корень проекта в sys.path
    _root = Path(__file__).resolve().parents[2]
    sys.path.insert(0, str(_root))

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFrame, QMessageBox, QGridLayout, QScrollArea,
)
from PyQt6.QtCore import Qt, pyqtSignal

from core import data_store


class DashboardPage(QWidget):
    """Главная страница с карточками действий."""

    open_history      = pyqtSignal()
    clear_last_routes = pyqtSignal()
    open_rounding_settings = pyqtSignal()

    def __init__(self, app_state: dict):
        super().__init__()
        self.app_state = app_state
        self._build_ui()

    # ─────────────────────────── UI ───────────────────────────────────

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
        inner.setContentsMargins(48, 40, 48, 40)
        inner.setSpacing(32)
        scroll.setWidget(content)

        # Заголовок
        title = QLabel("Маршруты, Сборка")
        title.setObjectName("sectionTitle")
        inner.addWidget(title)

        hint = QLabel(
            "Откройте историю сохранённых маршрутов. Обработка файлов и этикетки — через вкладки ленты или меню Файл."
        )
        hint.setObjectName("stepLabel")
        hint.setWordWrap(True)
        inner.addWidget(hint)

        grid = QGridLayout()
        grid.setSpacing(16)

        self._card_history = self._make_card(
            "📋", "История",
            "Открыть историю маршрутов (Основные и Увеличение)",
            "btnPrimary", self.open_history.emit
        )
        grid.addWidget(self._card_history, 0, 0)

        inner.addLayout(grid)

        # Нижняя панель с дополнительными действиями
        bottom_row = QHBoxLayout()
        bottom_row.addStretch()

        btn_rounding = QPushButton("Настройки Количества")
        btn_rounding.setObjectName("btnSecondary")
        btn_rounding.setMinimumWidth(260)
        btn_rounding.clicked.connect(self.open_rounding_settings.emit)
        bottom_row.addWidget(btn_rounding)

        btn_clear = QPushButton("Очистить историю")
        btn_clear.setObjectName("btnDanger")
        btn_clear.setMinimumWidth(220)
        btn_clear.clicked.connect(self._on_clear_last)
        bottom_row.addWidget(btn_clear)

        inner.addLayout(bottom_row)

        inner.addStretch()

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
        lbl_title.setWordWrap(True)
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
        btn.setMinimumWidth(180)
        btn.clicked.connect(on_click)
        card_lay.addWidget(btn)

        return card

    # ─────────────────────────── Обновление ───────────────────────────

    def refresh(self):
        """Обновляет подсказки и доступность карточки «История»."""
        main_hist = data_store.get_routes_history("main")
        inc_hist = data_store.get_routes_history("increase")
        total = len(main_hist) + len(inc_hist)

        if total > 0:
            tip = f"Основные: {len(main_hist)}, Увеличение: {len(inc_hist)}. Выберите сохранение из списка."
            self._card_history.setToolTip(tip)
            for w in self._card_history.findChildren(QPushButton):
                w.setEnabled(True)
                w.setToolTip(tip)
                break
        else:
            self._card_history.setToolTip("История пуста. Сначала обработайте файлы.")
            for w in self._card_history.findChildren(QPushButton):
                w.setEnabled(False)
                break

    def _on_clear_last(self):
        reply = QMessageBox.question(
            self, "Очистить историю",
            "Удалить всю историю маршрутов (основной и довоз)?\n"
            "После этого кнопки «История» не будут открывать данные.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.clear_last_routes.emit()
