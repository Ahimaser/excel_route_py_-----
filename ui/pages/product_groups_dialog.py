"""
product_groups_dialog.py — Диалог настройки групп продуктов для режима «разделить по продуктам».

Позволяет объединять несколько продуктов в один файл.
"""
from __future__ import annotations

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QListWidget,
    QListWidgetItem,
    QScrollArea,
    QFrame,
    QGroupBox,
    QMessageBox,
    QSplitter,
)

from core import data_store
from ui.widgets import hint_icon_button


def open_product_groups_dialog(
    parent, dept_key: str, dept_name: str, product_names: list[str],
    category: str | None = None,
) -> dict[str, list[list[str]]] | None:
    """
    Открывает диалог настройки групп продуктов.
    category: "ШК" | "СД" — для сохранения по категории (Школы/Сады). None — общий формат.
    Возвращает {dept_key: [[p1, p2], [p3], ...]} или None при отмене.
    """
    dlg = ProductGroupsDialog(parent, dept_key, dept_name, product_names, category=category)
    if dlg.exec() == QDialog.DialogCode.Accepted:
        return dlg.get_result()
    return None


class ProductGroupsDialog(QDialog):
    """Диалог настройки групп продуктов для одного отдела/подотдела."""

    def __init__(self, parent, dept_key: str, dept_name: str, product_names: list[str], category: str | None = None):
        super().__init__(parent)
        self.dept_key = dept_key
        self.dept_name = dept_name
        self.product_names = list(product_names or [])
        self.category = category
        cat_label = " (Школы)" if category == "ШК" else " (Сады)" if category == "СД" else ""
        self.setWindowTitle(f"Группы продуктов — {dept_name}{cat_label}")
        self.setMinimumSize(600, 450)
        self.resize(700, 520)

        saved_groups = data_store.get_product_file_groups(dept_key, category)
        if saved_groups:
            self._groups = [[p for p in g] for g in saved_groups]
        else:
            self._groups = [[p] for p in self.product_names]

        self._build_ui()

    def _build_ui(self) -> None:
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(16)

        hint_row = QHBoxLayout()
        hint_lbl = QLabel("Объедините продукты в группы. Каждая группа — отдельный файл.")
        hint_lbl.setObjectName("stepLabel")
        hint_lbl.setWordWrap(True)
        hint_row.addWidget(hint_lbl)
        hint_row.addWidget(hint_icon_button(
            self,
            "Группы определяют, как продукты объединяются в файлы при режиме «По группам».",
            "Инструкция — Группы продуктов\n\n"
            "1. Слева — список продуктов отдела. Справа — группы (каждая группа = отдельный файл).\n"
            "2. Выберите продукт слева, нажмите «+» у нужной группы — продукт добавится в группу.\n"
            "3. «Добавить группу» — новая пустая группа.\n"
            "4. «Удалить» — удалить группу (продукты останутся в отделе, но не войдут в файлы).",
            "Инструкция",
        ))
        hint_row.addStretch()
        lay.addLayout(hint_row)

        split = QSplitter(Qt.Orientation.Horizontal)

        left = QFrame()
        left_lay = QVBoxLayout(left)
        left_lay.addWidget(QLabel("Продукты отдела:"))
        self.products_list = QListWidget()
        self.products_list.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
        for p in self.product_names:
            self.products_list.addItem(p)
        left_lay.addWidget(self.products_list)
        split.addWidget(left)

        right = QFrame()
        right_lay = QVBoxLayout(right)
        right_lay.addWidget(QLabel("Группы (файл на группу):"))
        self.groups_scroll = QScrollArea()
        self.groups_scroll.setWidgetResizable(True)
        self.groups_widget = QFrame()
        self.groups_lay = QVBoxLayout(self.groups_widget)
        self.groups_lay.setContentsMargins(0, 0, 0, 0)
        self._group_widgets: list[tuple[QFrame, QListWidget]] = []
        self._refresh_groups_ui()
        self.groups_scroll.setWidget(self.groups_widget)
        right_lay.addWidget(self.groups_scroll)

        btn_row = QHBoxLayout()
        self.btn_add_group = QPushButton("Добавить группу")
        self.btn_add_group.clicked.connect(self._on_add_group)
        btn_row.addWidget(self.btn_add_group)
        btn_row.addStretch()
        right_lay.addLayout(btn_row)

        split.addWidget(right)
        split.setSizes([200, 320])
        lay.addWidget(split)

        btn_row_main = QHBoxLayout()
        btn_row_main.addStretch()
        btn_ok = QPushButton("OK")
        btn_ok.setObjectName("btnPrimary")
        btn_ok.clicked.connect(self.accept)
        btn_cancel = QPushButton("Отмена")
        btn_cancel.setObjectName("btnSecondary")
        btn_cancel.clicked.connect(self.reject)
        btn_row_main.addWidget(btn_ok)
        btn_row_main.addWidget(btn_cancel)
        lay.addLayout(btn_row_main)

    def _refresh_groups_ui(self) -> None:
        for w, _ in self._group_widgets:
            w.deleteLater()
        self._group_widgets.clear()

        for gi, group in enumerate(self._groups):
            grp_frame = QFrame()
            grp_frame.setObjectName("card")
            grp_lay = QHBoxLayout(grp_frame)
            grp_lay.setContentsMargins(8, 4, 8, 4)

            lst = QListWidget()
            lst.setMaximumHeight(80)
            lst.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
            for p in group:
                lst.addItem(p)
            grp_lay.addWidget(lst, 1)

            btn_add = QPushButton("+")
            btn_add.setToolTip("Добавить выбранный продукт в группу")
            btn_add.setProperty("group_idx", gi)
            btn_add.clicked.connect(self._on_add_to_group)
            grp_lay.addWidget(btn_add)

            btn_remove = QPushButton("Удалить")
            btn_remove.setToolTip("Удалить группу")
            btn_remove.setProperty("group_idx", gi)
            btn_remove.clicked.connect(self._on_remove_group)
            grp_lay.addWidget(btn_remove)

            self._group_widgets.append((grp_frame, lst))
            self.groups_lay.addWidget(grp_frame)

    def _on_add_group(self) -> None:
        self._groups.append([])
        self._refresh_groups_ui()

    def _on_add_to_group(self) -> None:
        row = self.products_list.currentRow()
        if row < 0:
            return
        prod = self.products_list.item(row).text()
        btn = self.sender()
        if btn is None:
            return
        gi = btn.property("group_idx")
        if gi is None or gi < 0 or gi >= len(self._groups):
            return
        if prod not in self._groups[gi]:
            self._groups[gi].append(prod)
            self._refresh_groups_ui()

    def _on_remove_group(self) -> None:
        btn = self.sender()
        if btn and hasattr(btn, "property"):
            gi = btn.property("group_idx")
            if gi is not None and 0 <= gi < len(self._groups):
                self._groups.pop(gi)
                self._refresh_groups_ui()

    def _collect_groups_from_ui(self) -> list[list[str]]:
        result = []
        for _, lst in self._group_widgets:
            items = [lst.item(i).text() for i in range(lst.count())]
            if items:
                result.append(items)
        return result

    def get_result(self) -> dict[str, list[list[str]]]:
        groups = self._collect_groups_from_ui()
        if not groups:
            groups = [[p] for p in self.product_names]
        data_store.set_product_file_groups(self.dept_key, groups, category=self.category)
        return {self.dept_key: groups}
