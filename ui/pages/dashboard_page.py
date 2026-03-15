"""
dashboard_page.py — Главная страница приложения (дашборд).

Описание программы, место сохранения, отчёт по последним маршрутам.
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


def _routes_summary(blob: dict | None) -> tuple[int, int, list[str]]:
    """Возвращает (кол-во маршрутов, кол-во продуктов, список названий продуктов)."""
    if not blob:
        return 0, 0, []
    routes = blob.get("filteredRoutes") or blob.get("routes") or []
    active = [r for r in routes if not r.get("excluded")]
    products_set: set[str] = set()
    for r in active:
        for p in r.get("products", []):
            name = p.get("name")
            if name:
                products_set.add(name)
    return len(active), len(products_set), sorted(products_set)


class DashboardPage(QWidget):
    """Главная страница: описание, место сохранения, отчёт по маршрутам."""

    open_history = pyqtSignal()
    go_last_main = pyqtSignal()
    go_last_increase = pyqtSignal()

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
        inner.setSpacing(20)
        scroll.setWidget(content)

        title = QLabel("Маршруты, Сборка")
        title.setObjectName("sectionTitle")
        inner.addWidget(title)

        # Краткое описание программы
        desc = QLabel(
            "Приложение для обработки маршрутных XLS-файлов (школы и сады), "
            "генерации отчётов по отделам и создания этикеток. "
            "Загрузка файлов, предпросмотр и выгрузка — через вкладки ленты сверху."
        )
        desc.setObjectName("stepLabel")
        desc.setWordWrap(True)
        inner.addWidget(desc)

        # Место сохранения маршрутов
        save_frame = QFrame()
        save_frame.setObjectName("card")
        save_lay = QVBoxLayout(save_frame)
        save_lay.setContentsMargins(16, 12, 16, 12)
        save_lay.setSpacing(6)
        save_lay.addWidget(QLabel("Папка сохранения маршрутов:"))
        self.lbl_save_dir = QLabel("")
        self.lbl_save_dir.setObjectName("hintLabel")
        self.lbl_save_dir.setWordWrap(True)
        save_lay.addWidget(self.lbl_save_dir)
        inner.addWidget(save_frame)

        # Отчёт по последним маршрутам
        report_frame = QFrame()
        report_frame.setObjectName("card")
        report_lay = QVBoxLayout(report_frame)
        report_lay.setContentsMargins(16, 12, 16, 12)
        report_lay.setSpacing(12)
        report_lay.addWidget(QLabel("Последние маршруты"))
        report_lay.addWidget(QLabel("Краткий отчёт по сохранённым маршрутам."))
        self.lbl_report_main = QLabel("")
        self.lbl_report_main.setObjectName("hintLabel")
        self.lbl_report_main.setWordWrap(True)
        self.lbl_report_main.setTextFormat(Qt.TextFormat.RichText)
        report_lay.addWidget(self.lbl_report_main)
        self.lbl_report_inc = QLabel("")
        self.lbl_report_inc.setObjectName("hintLabel")
        self.lbl_report_inc.setWordWrap(True)
        self.lbl_report_inc.setTextFormat(Qt.TextFormat.RichText)
        report_lay.addWidget(self.lbl_report_inc)
        inner.addWidget(report_frame)

        # Кнопки
        grid = QGridLayout()
        grid.setSpacing(10)

        self._card_history = self._make_card(
            "📋", "История",
            "Выбор сохранённых маршрутов из списка (основные и довоз)",
            "btnPrimary", self.open_history.emit, "history"
        )
        grid.addWidget(self._card_history, 0, 0)

        self._card_last_main = self._make_card(
            "📄", "Последние (основной)",
            "Открыть последние сохранённые маршруты основного типа",
            "btnSecondary", self.go_last_main.emit, "last_main"
        )
        grid.addWidget(self._card_last_main, 0, 1)

        self._card_last_inc = self._make_card(
            "📄", "Последние (довоз)",
            "Открыть последние сохранённые маршруты довоза",
            "btnSecondary", self.go_last_increase.emit, "last_inc"
        )
        grid.addWidget(self._card_last_inc, 0, 2)

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
        lbl_icon.setObjectName("reportIcon")
        lbl_icon.setAlignment(Qt.AlignmentFlag.AlignLeft)
        card_lay.addWidget(lbl_icon)

        lbl_title = QLabel(title)
        lbl_title.setObjectName("cardTitle")
        lbl_title.setWordWrap(True)
        card_lay.addWidget(lbl_title)

        lbl_desc = QLabel(desc)
        lbl_desc.setObjectName("reportDesc")
        lbl_desc.setWordWrap(True)
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
        """Обновляет место сохранения, отчёт и доступность карточек."""
        save_dir = (
            data_store.get_setting("defaultSaveDir") or
            data_store.get_desktop_path()
        )
        self.lbl_save_dir.setText(f"📁 {save_dir}")

        main_data = data_store.get_last_routes("main")
        inc_data = data_store.get_last_routes("increase")

        n_main, p_main, prods_main = _routes_summary(main_data)
        n_inc, p_inc, prods_inc = _routes_summary(inc_data)

        if main_data:
            prods_preview = ", ".join(prods_main[:8])
            if len(prods_main) > 8:
                prods_preview += f" … (+{len(prods_main) - 8})"
            self.lbl_report_main.setText(
                f"<b>Основной:</b> {n_main} маршрутов, {p_main} продуктов. "
                f"Продукты: {prods_preview or '—'}"
            )
        else:
            self.lbl_report_main.setText("<b>Основной:</b> нет данных")

        if inc_data:
            prods_preview = ", ".join(prods_inc[:8])
            if len(prods_inc) > 8:
                prods_preview += f" … (+{len(prods_inc) - 8})"
            self.lbl_report_inc.setText(
                f"<b>Довоз:</b> {n_inc} маршрутов, {p_inc} продуктов. "
                f"Продукты: {prods_preview or '—'}"
            )
        else:
            self.lbl_report_inc.setText("<b>Довоз:</b> нет данных")

        # История
        main_hist = data_store.get_routes_history("main")
        inc_hist = data_store.get_routes_history("increase")
        total = len(main_hist) + len(inc_hist)
        if total > 0:
            tip = f"Основные: {len(main_hist)}, Довоз: {len(inc_hist)}. Выберите сохранение из списка."
            self._set_card_enabled(self._card_history, True, tip)
        else:
            self._set_card_enabled(self._card_history, False, "История пуста. Сначала обработайте файлы.")

        has_main = n_main > 0
        has_inc = n_inc > 0
        self._set_card_enabled(
            self._card_last_main,
            has_main,
            "Открыть последние маршруты основного типа" if has_main else "Нет сохранённых маршрутов. Сначала обработайте файлы."
        )
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
