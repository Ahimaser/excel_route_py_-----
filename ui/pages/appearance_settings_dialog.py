"""
appearance_settings_dialog.py — Настройки оформления: тёмная тема.
"""
from __future__ import annotations

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QCheckBox,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QShortcut, QKeySequence

from core import data_store


class AppearanceSettingsDialog(QDialog):
    """Диалог настроек оформления: тёмная тема."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Оформление")
        self.setMinimumWidth(360)
        lay = QVBoxLayout(self)

        self.chk_dark = QCheckBox("Тёмная тема")
        self.chk_dark.setChecked(bool(data_store.get_setting("darkTheme")))
        self.chk_dark.setToolTip("Переключение между светлой и тёмной темой интерфейса")
        lay.addWidget(self.chk_dark)

        hint = QLabel(
            "Для применения темы потребуется перезапуск приложения."
        )
        hint.setObjectName("hintLabel")
        hint.setWordWrap(True)
        lay.addWidget(hint)

        btn_lay = QHBoxLayout()
        btn_lay.addStretch()
        btn_ok = QPushButton("Сохранить")
        btn_ok.setObjectName("btnPrimary")
        btn_ok.setDefault(True)
        btn_ok.clicked.connect(self._save)
        btn_lay.addWidget(btn_ok)
        lay.addLayout(btn_lay)
        QShortcut(QKeySequence(Qt.Key.Key_Return), self, self._save)

    def _save(self):
        data_store.set_setting("darkTheme", self.chk_dark.isChecked())
        self.accept()


def open_appearance_settings_dialog(parent=None):
    dlg = AppearanceSettingsDialog(parent)
    dlg.exec()
