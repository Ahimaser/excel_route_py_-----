"""
add_product_dialog.py — Диалог добавления продукта (карточная форма).

Вариант A из ADD_PRODUCT_UI_PROPOSALS: все поля на одном экране.
"""
from __future__ import annotations

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QShortcut, QKeySequence
from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QComboBox,
    QPushButton,
    QFrame,
    QCheckBox,
    QDoubleSpinBox,
    QWidget,
)

from core import data_store

COMMON_UNITS = ["кг", "л", "шт", "г", "мл", "уп", "бан", "пак"]


def open_add_product_dialog(parent: QWidget) -> bool:
    """
    Открывает диалог добавления продукта.
    Возвращает True, если продукт добавлен; False при отмене.
    """
    dlg = AddProductDialog(parent)
    return dlg.exec() == QDialog.DialogCode.Accepted


class AddProductDialog(QDialog):
    """Карточная форма добавления продукта в справочник."""

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Добавить продукт")
        self.setModal(True)
        self.setMinimumSize(420, 340)
        self.resize(460, 380)
        self._build_ui()

    def _build_ui(self) -> None:
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(16)

        card = QFrame()
        card.setObjectName("card")
        card_lay = QVBoxLayout(card)
        card_lay.setContentsMargins(20, 20, 20, 20)
        card_lay.setSpacing(14)

        # Название
        card_lay.addWidget(QLabel("Название продукта"))
        self.edit_name = QLineEdit()
        self.edit_name.setPlaceholderText("Масло сливочное")
        self.edit_name.setClearButtonEnabled(True)
        card_lay.addWidget(self.edit_name)

        # Единица измерения
        card_lay.addWidget(QLabel("Единица измерения"))
        self.combo_unit = QComboBox()
        self.combo_unit.setEditable(True)
        for u in COMMON_UNITS:
            self.combo_unit.addItem(u)
        self.combo_unit.setCurrentText("кг")
        self.combo_unit.currentTextChanged.connect(self._on_unit_changed)
        card_lay.addWidget(self.combo_unit)

        # Отдел
        card_lay.addWidget(QLabel("Отдел"))
        self.combo_dept = QComboBox()
        self.combo_dept.addItem("Без отдела", "")
        for key, name in data_store.get_department_choices():
            if key:
                self.combo_dept.addItem(name, key)
        card_lay.addWidget(self.combo_dept)

        # Вариант (кол-во в 1 шт)
        self.chk_variant = QCheckBox("Создать вариант (кол-во в 1 шт/коробке)")
        self.chk_variant.setToolTip("Создаст продукт вида «Название (0,18 кг)» с настройкой Шт")
        self.chk_variant.toggled.connect(self._on_variant_toggled)
        card_lay.addWidget(self.chk_variant)

        variant_row = QHBoxLayout()
        variant_row.addWidget(QLabel("Кол-во в 1 шт:"))
        self.spin_pcu = QDoubleSpinBox()
        self.spin_pcu.setRange(0.001, 999.999)
        self.spin_pcu.setDecimals(3)
        self.spin_pcu.setSingleStep(0.1)
        self.spin_pcu.setValue(0.18)
        self.spin_pcu.setFixedWidth(100)
        variant_row.addWidget(self.spin_pcu)
        self.lbl_unit_suffix = QLabel("кг")
        self.lbl_unit_suffix.setObjectName("unitLabel")
        variant_row.addWidget(self.lbl_unit_suffix)
        variant_row.addStretch()
        self._variant_row_widget = QWidget()
        self._variant_row_widget.setLayout(variant_row)
        self._variant_row_widget.setVisible(False)
        card_lay.addWidget(self._variant_row_widget)

        lay.addWidget(card)

        # Кнопки
        btn_lay = QHBoxLayout()
        btn_lay.addStretch()
        btn_cancel = QPushButton("Отмена")
        btn_cancel.setObjectName("btnSecondary")
        btn_cancel.clicked.connect(self.reject)
        btn_add = QPushButton("Добавить")
        btn_add.setObjectName("btnPrimary")
        btn_add.setDefault(True)
        btn_add.setAutoDefault(True)
        btn_add.clicked.connect(self._on_add_clicked)
        btn_lay.addWidget(btn_cancel)
        btn_lay.addWidget(btn_add)
        lay.addLayout(btn_lay)

        QShortcut(QKeySequence(Qt.Key.Key_Escape), self, self.reject)
        QShortcut(QKeySequence(Qt.Key.Key_Return), self, self._on_add_clicked)

        self.edit_name.setFocus()

    def _on_unit_changed(self, text: str) -> None:
        self.lbl_unit_suffix.setText((text or "").strip() or "—")
        unit = (text or "").strip().lower()
        can_variant = unit and unit != "шт"
        self.chk_variant.setEnabled(can_variant)
        if not can_variant:
            self.chk_variant.setChecked(False)

    def _on_variant_toggled(self, checked: bool) -> None:
        self._variant_row_widget.setVisible(checked)

    def _on_add_clicked(self) -> None:
        name = (self.edit_name.text() or "").strip()
        if not name:
            self.edit_name.setFocus()
            return
        unit = (self.combo_unit.currentText() or "").strip()
        dept_key = self.combo_dept.currentData() or ""
        as_variant = self.chk_variant.isChecked() and unit and unit.lower() != "шт"

        products = data_store.get_ref("products") or []
        if as_variant:
            pcu = self.spin_pcu.value()
            pcu_str = str(pcu).replace(".", ",")
            final_name = f"{name} ({pcu_str} {unit})"
        else:
            final_name = name

        if any(p.get("name") == final_name for p in products):
            from PyQt6.QtWidgets import QMessageBox
            QMessageBox.warning(
                self, "Ошибка",
                f"Продукт «{final_name}» уже есть в справочнике.",
            )
            return

        if data_store.add_product(final_name, unit, dept_key or None):
            if as_variant:
                data_store.update_product(final_name, showPcs=True, pcsPerUnit=self.spin_pcu.value())
            self.accept()
        else:
            from PyQt6.QtWidgets import QMessageBox
            QMessageBox.warning(
                self, "Ошибка",
                f"Не удалось добавить «{final_name}» (возможно, уже существует).",
            )
