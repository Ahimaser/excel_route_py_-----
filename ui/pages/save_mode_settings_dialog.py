"""
save_mode_settings_dialog.py — Окно настройки режимов сохранения по отделам для ШК и СД.

Вариант A: таблица-матрица.
Строки — отделы и подотделы, столбцы — ШК (школы) и СД (сады).
В каждой ячейке выбор режима: все в один файл / по группам / по продуктам.
"""
from __future__ import annotations

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QComboBox,
    QAbstractItemView,
    QMessageBox,
    QMenu,
    QWidget,
)

from core import data_store
from ui.widgets import hint_icon_button

SAVE_MODE_ALL = "all"
SAVE_MODE_GROUPS = "groups"
SAVE_MODE_BY_PRODUCT = "by_product"

MODE_LABELS = {
    SAVE_MODE_ALL: "Все в один файл",
    SAVE_MODE_GROUPS: "По группам",
    SAVE_MODE_BY_PRODUCT: "По продуктам",
}


def open_save_mode_settings_dialog(parent: QWidget | None) -> None:
    """Открывает диалог настройки режимов сохранения для ШК и СД."""
    dlg = SaveModeSettingsDialog(parent)
    dlg.exec()


class SaveModeSettingsDialog(QDialog):
    """Диалог: таблица-матрица отделы × (ШК, СД) с выбором режима сохранения."""

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Режимы сохранения — Школы и Сады")
        self.setMinimumSize(700, 500)
        self.resize(800, 550)
        self._build_ui()
        self._load_data()

    def _build_ui(self) -> None:
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(16)

        hint_row = QHBoxLayout()
        hint = QLabel(
            "Режим сохранения файлов по отделам для маршрутов ШК (школы) и СД (сады). "
            "«Все в один» — один файл на отдел; «По группам» — файл на группу продуктов; "
            "«По продуктам» — файл на каждый продукт."
        )
        hint.setObjectName("stepLabel")
        hint.setWordWrap(True)
        hint_row.addWidget(hint)
        hint_row.addWidget(hint_icon_button(
            self,
            "Режим сохранения для каждого отдела и категории (ШК/СД).",
            "Инструкция — Режимы сохранения\n\n"
            "1. «Все в один файл» — один Excel-файл на отдел для выбранной категории (ШК или СД).\n"
            "2. «По группам» — отдельный файл на каждую группу продуктов. Группы настраиваются кнопкой «Настроить…».\n"
            "3. «По продуктам» — отдельный файл на каждый продукт отдела.\n"
            "4. Кнопка «Настроить…» доступна только при режиме «По группам» для соответствующей категории.",
            "Инструкция",
        ))
        hint_row.addStretch()
        lay.addLayout(hint_row)

        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["Отдел / подотдел", "ШК (школы)", "СД (сады)", "Группы"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        lay.addWidget(self.table, 1)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_ok = QPushButton("OK")
        btn_ok.setObjectName("btnPrimary")
        btn_ok.clicked.connect(self._on_save)
        btn_cancel = QPushButton("Отмена")
        btn_cancel.setObjectName("btnSecondary")
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_ok)
        btn_row.addWidget(btn_cancel)
        lay.addLayout(btn_row)

    def _load_data(self) -> None:
        """Заполняет таблицу отделами и подотделами из data_store."""
        depts = data_store.get_ref("departments") or []
        products = data_store.get_ref("products") or []
        prod_by_dept: dict[str, list[str]] = {}
        for p in products:
            k = p.get("deptKey")
            if k:
                prod_by_dept.setdefault(k, []).append(p.get("name", ""))

        rows: list[tuple[str, str, bool]] = []  # (display_name, dept_key, is_subdept)
        for dept in depts:
            for sub in dept.get("subdepts", []):
                sub_key = sub.get("key", "")
                if prod_by_dept.get(sub_key):
                    rows.append((f"  └ {sub.get('name', '')}", sub_key, True))
            dept_key = dept.get("key", "")
            if prod_by_dept.get(dept_key):
                rows.append((dept.get("name", ""), dept_key, False))

        self.table.setRowCount(len(rows))
        self._row_data: list[tuple[str, bool]] = []  # (dept_key, is_subdept)

        for i, (display_name, dept_key, is_subdept) in enumerate(rows):
            self._row_data.append((dept_key, is_subdept))
            name_item = QTableWidgetItem(display_name)
            name_item.setFlags(name_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(i, 0, name_item)

            for col, cat in enumerate(("ШК", "СД"), start=1):
                combo = QComboBox()
                for val, lbl in MODE_LABELS.items():
                    combo.addItem(lbl, val)
                mode = data_store.get_save_mode_by_dept_category(dept_key, cat)
                idx = combo.findData(mode)
                combo.setCurrentIndex(idx if idx >= 0 else 0)
                combo.setProperty("dept_key", dept_key)
                combo.setProperty("category", cat)
                combo.currentIndexChanged.connect(self._on_mode_changed)
                self.table.setCellWidget(i, col, combo)

            btn = QPushButton("Настроить…")
            btn.setProperty("dept_key", dept_key)
            btn.setProperty("row", i)
            btn.clicked.connect(self._on_config_groups)
            self.table.setCellWidget(i, 3, btn)

        self._update_buttons()

    def _update_buttons(self) -> None:
        """Включает кнопку «Настроить группы» только при режиме «По группам»."""
        for i in range(self.table.rowCount()):
            btn = self.table.cellWidget(i, 3)
            if btn:
                shk_combo = self.table.cellWidget(i, 1)
                sd_combo = self.table.cellWidget(i, 2)
                shk_mode = shk_combo.currentData() if shk_combo else SAVE_MODE_ALL
                sd_mode = sd_combo.currentData() if sd_combo else SAVE_MODE_ALL
                needs_config = shk_mode == SAVE_MODE_GROUPS or sd_mode == SAVE_MODE_GROUPS
                btn.setEnabled(needs_config)

    def _on_mode_changed(self) -> None:
        """Обновляет доступность кнопки «Настроить группы» при смене режима."""
        self._update_buttons()

    def _on_config_groups(self) -> None:
        """Открывает меню выбора категории и диалог групп продуктов."""
        btn = self.sender()
        if not btn:
            return
        dept_key = btn.property("dept_key")
        row = btn.property("row")
        if row is None:
            row = -1
        if row < 0:
            return

        dept_name = self.table.item(row, 0).text().strip().lstrip("└ ")
        products = data_store.get_ref("products") or []
        product_names = [p["name"] for p in products if p.get("deptKey") == dept_key]
        if not product_names:
            QMessageBox.information(
                self, "Нет продуктов",
                f"У отдела «{dept_name}» нет привязанных продуктов."
            )
            return

        menu = QMenu(self)
        act_shk = menu.addAction("ШК (школы)")
        act_sd = menu.addAction("СД (сады)")
        act = menu.exec(btn.mapToGlobal(btn.rect().bottomLeft()))
        if act == act_shk:
            self._open_product_groups(dept_key, dept_name, product_names, "ШК")
        elif act == act_sd:
            self._open_product_groups(dept_key, dept_name, product_names, "СД")

    def _open_product_groups(
        self, dept_key: str, dept_name: str, product_names: list[str], category: str
    ) -> None:
        from ui.pages.product_groups_dialog import open_product_groups_dialog
        open_product_groups_dialog(self, dept_key, dept_name, product_names, category=category)

    def _on_save(self) -> None:
        """Сохраняет выбранные режимы в data_store."""
        for i in range(self.table.rowCount()):
            if i >= len(self._row_data):
                continue
            dept_key, _ = self._row_data[i]
            for col, cat in enumerate(("ШК", "СД"), start=1):
                combo = self.table.cellWidget(i, col)
                if combo:
                    mode = combo.currentData() or SAVE_MODE_ALL
                    data_store.set_save_mode_by_dept_category(dept_key, cat, mode)
        self.accept()
