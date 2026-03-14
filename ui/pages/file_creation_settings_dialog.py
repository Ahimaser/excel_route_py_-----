"""
file_creation_settings_dialog.py — Настройки по умолчанию для создания файлов: размер шрифта, отступы страницы.
К этикеткам не применяются.
"""
from __future__ import annotations

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFormLayout, QSpinBox, QDoubleSpinBox, QGroupBox,
)
from PyQt6.QtCore import Qt

from core import data_store


class FileCreationSettingsDialog(QDialog):
    """Диалог: размер шрифта и отступы страницы для создаваемых XLS (кроме этикеток)."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Параметры создания файлов")
        self.setMinimumWidth(360)
        lay = QVBoxLayout(self)

        gr = QGroupBox("По умолчанию для файлов маршрутов (не для этикеток)")
        form = QFormLayout(gr)
        self.spin_font = QSpinBox()
        self.spin_font.setRange(8, 24)
        self.spin_font.setValue(int(data_store.get_setting("defaultFontSize") or 12))
        self.spin_font.setSuffix(" pt")
        form.addRow("Размер текста:", self.spin_font)

        self.spin_top = QDoubleSpinBox()
        self.spin_top.setRange(0.2, 5.0)
        self.spin_top.setDecimals(1)
        self.spin_top.setValue(float(data_store.get_setting("defaultMarginTop") or 1.5))
        self.spin_top.setSuffix(" см")
        form.addRow("Отступ сверху:", self.spin_top)
        self.spin_left = QDoubleSpinBox()
        self.spin_left.setRange(0.2, 5.0)
        self.spin_left.setDecimals(1)
        self.spin_left.setValue(float(data_store.get_setting("defaultMarginLeft") or 1.5))
        self.spin_left.setSuffix(" см")
        form.addRow("Отступ слева:", self.spin_left)
        self.spin_bottom = QDoubleSpinBox()
        self.spin_bottom.setRange(0.1, 5.0)
        self.spin_bottom.setDecimals(1)
        self.spin_bottom.setValue(float(data_store.get_setting("defaultMarginBottom") or 0.5))
        self.spin_bottom.setSuffix(" см")
        form.addRow("Отступ снизу:", self.spin_bottom)
        self.spin_right = QDoubleSpinBox()
        self.spin_right.setRange(0.1, 5.0)
        self.spin_right.setDecimals(1)
        self.spin_right.setValue(float(data_store.get_setting("defaultMarginRight") or 0.5))
        self.spin_right.setSuffix(" см")
        form.addRow("Отступ справа:", self.spin_right)

        lay.addWidget(gr)
        hint = QLabel("Эти настройки применяются к файлам маршрутов (общие и по отделам). Этикетки не изменяются.")
        hint.setWordWrap(True)
        hint.setObjectName("hintLabel")
        lay.addWidget(hint)

        btn_lay = QHBoxLayout()
        btn_lay.addStretch()
        btn_ok = QPushButton("Сохранить")
        btn_ok.setObjectName("btnPrimary")
        btn_ok.clicked.connect(self._save)
        btn_lay.addWidget(btn_ok)
        lay.addLayout(btn_lay)

    def _save(self):
        data_store.set_setting("defaultFontSize", self.spin_font.value())
        data_store.set_setting("defaultMarginTop", self.spin_top.value())
        data_store.set_setting("defaultMarginLeft", self.spin_left.value())
        data_store.set_setting("defaultMarginBottom", self.spin_bottom.value())
        data_store.set_setting("defaultMarginRight", self.spin_right.value())
        try:
            from core import excel_generator
            excel_generator._STYLES = None
            excel_generator._STYLES_FONT_PT = None
        except Exception:
            pass
        self.accept()


def open_file_creation_settings_dialog(parent=None):
    dlg = FileCreationSettingsDialog(parent)
    dlg.exec()
