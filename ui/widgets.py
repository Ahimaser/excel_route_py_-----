"""
ui/widgets.py — Общие переиспользуемые виджеты.

CommitLineEdit — поле ввода, которое надёжно сохраняет изменения:
  - при потере фокуса (клик в другое место экрана)
  - при нажатии Enter
  - НЕ срабатывает дважды (флаг _committed)
  - Безопасно при удалении виджета: текст сохраняется до emit

Архитектура:
  Вместо emit() прямо в focusOutEvent используем pending_value —
  сохраняем текст и испускаем сигнал через QTimer.singleShot(0).
  Это гарантирует, что Qt завершит обработку события фокуса
  до того, как слот может удалить виджет.
"""
from __future__ import annotations

from PyQt6.QtWidgets import QLineEdit
from PyQt6.QtCore import pyqtSignal, QTimer


class CommitLineEdit(QLineEdit):
    """Поле ввода с надёжным сохранением по Enter и потере фокуса.

    Использование:
        editor = CommitLineEdit("начальный текст")
        editor.commit.connect(lambda: save(editor.pending_value))
        # или просто:
        editor.commit.connect(lambda: save(editor.text()))
    """
    commit = pyqtSignal()

    def __init__(self, text: str = "", parent=None):
        super().__init__(text, parent)
        self._committed = False
        self.pending_value: str = text  # текст на момент commit, безопасен после удаления
        self.returnPressed.connect(self._schedule_commit)

    def reset_commit(self):
        """Сбросить флаг, чтобы поле снова могло испустить commit."""
        self._committed = False
        self.pending_value = self.text()

    def focusOutEvent(self, event):
        """Потеря фокуса → запланировать сохранение."""
        self._schedule_commit()
        super().focusOutEvent(event)

    def _schedule_commit(self):
        """Сохранить текст и запланировать emit через event loop."""
        if not self._committed:
            self._committed = True
            self.pending_value = self.text()  # сохраняем ДО удаления виджета
            QTimer.singleShot(0, self._safe_emit)

    def _safe_emit(self):
        """Испускаем сигнал — виджет может быть уже удалён, но pending_value цел."""
        try:
            self.commit.emit()
        except RuntimeError:
            # C++ объект уже удалён — игнорируем
            pass
