"""
license_check.py — Проверка лицензии для скомпилированного exe.

Логика (только при getattr(sys, "frozen", False)):
- Если текущая дата >= 1 июня 2026: программа не запускается, показывается сообщение
  «Вы не оплатили программу».
- До этой даты — работа без ограничений.
"""
from __future__ import annotations

import sys
from datetime import date

# Дата блокировки: с этой даты (включительно) exe не запускается
_BLOCK_DATE = date(2026, 6, 1)


def _is_frozen() -> bool:
    """True если приложение запущено как скомпилированный exe."""
    return getattr(sys, "frozen", False)


def check_license(app=None) -> bool:
    """
    Проверяет лицензию. Возвращает True если запуск разрешён.
    Только для скомпилированного exe; при python app.py всегда True.
    app — QApplication (для отображения диалога).
    """
    if not _is_frozen():
        return True

    today = date.today()
    if today < _BLOCK_DATE:
        return True

    if app:
        from PyQt6.QtWidgets import QMessageBox
        QMessageBox.critical(
            None,
            "Доступ заблокирован",
            "Вы не оплатили программу",
        )
    return False
