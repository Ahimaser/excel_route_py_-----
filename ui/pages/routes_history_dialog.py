from __future__ import annotations

from copy import deepcopy

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QSplitter,
    QFrame,
    QWidget,
    QMessageBox,
    QMenu,
)

from core import data_store


def _type_title(file_type: str) -> str:
    return "Основной" if file_type == "main" else "Увеличение"


def _format_dt(iso_str: str) -> str:
    """Форматирует ISO дату-время в DD.MM.YYYY HH:MM."""
    if not iso_str:
        return ""
    s = str(iso_str).replace("T", " ")[:16]
    parts = s.split(" ")
    if len(parts) == 2:
        ymd = parts[0].split("-")
        if len(ymd) == 3:
            return f"{ymd[2]}.{ymd[1]}.{ymd[0]} {parts[1][:5]}"
    return s


def _entry_text(entry: dict) -> str:
    count = entry.get("count", 0)
    date_str = entry.get("date") or (entry.get("timestamp") or "")[:10]
    typ = _type_title(entry.get("fileType", "main"))
    cat = entry.get("routeCategory") or "ШК"
    created = _format_dt(entry.get("createdAt") or entry.get("timestamp"))
    modified = _format_dt(entry.get("modifiedAt") or entry.get("timestamp"))
    if date_str:
        parts = date_str.split("-")
        if len(parts) == 3:
            date_display = f"{parts[2]}.{parts[1]}.{parts[0]}"
        else:
            date_display = date_str
    else:
        date_display = created[:10] if created else ""
    if created and modified and created != modified:
        return f"{date_display} | {typ} | Созд. {created} / Изм. {modified} | {cat} | {count} марш."
    return f"{date_display} | {typ} | {cat} | маршрутов: {count}"


class RoutesHistoryDialog(QDialog):
    """Выбор записи из истории. Две колонки: Основные и Увеличение."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._history_main: list[dict] = []
        self._history_increase: list[dict] = []
        self._selected: dict | None = None
        self._loaded_entry: dict | None = None
        self._last_active_list: QListWidget | None = None
        self._build_ui()
        self._load_history()

    def _build_ui(self) -> None:
        self.setWindowTitle("История маршрутов")
        self.setMinimumSize(900, 560)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 16, 20, 16)
        lay.setSpacing(10)

        lbl = QLabel("Выберите сохранение из истории (двойной клик — открыть):")
        lbl.setObjectName("sectionTitle")
        lay.addWidget(lbl)

        self.search = QLineEdit()
        self.search.setPlaceholderText("Поиск по дате, типу или количеству маршрутов...")
        self.search.textChanged.connect(self._apply_filter)
        lay.addWidget(self.search)

        split = QSplitter(Qt.Orientation.Horizontal)

        # Колонка Основные
        col_main = QFrame()
        col_main.setObjectName("card")
        col_main_lay = QVBoxLayout(col_main)
        col_main_lay.setContentsMargins(12, 12, 12, 12)
        lbl_main = QLabel("Основные")
        lbl_main.setObjectName("sectionTitle")
        col_main_lay.addWidget(lbl_main)
        self.list_main = QListWidget()
        self.list_main.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.list_main.customContextMenuRequested.connect(self._on_context_menu)
        self.list_main.itemDoubleClicked.connect(lambda: self._accept_from_list(self.list_main))
        self.list_main.currentItemChanged.connect(lambda c, p, l=self.list_main: self._on_current_changed(c, l))
        col_main_lay.addWidget(self.list_main, 1)
        split.addWidget(col_main)

        # Колонка Увеличение
        col_inc = QFrame()
        col_inc.setObjectName("card")
        col_inc_lay = QVBoxLayout(col_inc)
        col_inc_lay.setContentsMargins(12, 12, 12, 12)
        lbl_inc = QLabel("Увеличение")
        lbl_inc.setObjectName("sectionTitle")
        col_inc_lay.addWidget(lbl_inc)
        self.list_increase = QListWidget()
        self.list_increase.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.list_increase.customContextMenuRequested.connect(self._on_context_menu)
        self.list_increase.itemDoubleClicked.connect(lambda: self._accept_from_list(self.list_increase))
        self.list_increase.currentItemChanged.connect(lambda c, p, l=self.list_increase: self._on_current_changed(c, l))
        col_inc_lay.addWidget(self.list_increase, 1)
        split.addWidget(col_inc)

        split.setSizes([450, 450])
        lay.addWidget(split, 1)

        self.lbl_meta = QLabel("")
        self.lbl_meta.setObjectName("hintLabel")
        self.lbl_meta.setWordWrap(True)
        lay.addWidget(self.lbl_meta)

        row = QHBoxLayout()
        row.addStretch()
        self.btn_open = QPushButton("Открыть")
        self.btn_open.setObjectName("btnPrimary")
        self.btn_open.setEnabled(False)
        self.btn_open.clicked.connect(self._on_open_clicked)
        row.addWidget(self.btn_open)
        btn_clear = QPushButton("Очистить историю")
        btn_clear.setObjectName("btnDanger")
        btn_clear.clicked.connect(self._on_clear_history)
        row.addWidget(btn_clear)
        btn_cancel = QPushButton("Отмена")
        btn_cancel.setObjectName("btnSecondary")
        btn_cancel.clicked.connect(self.reject)
        row.addWidget(btn_cancel)
        lay.addLayout(row)

    def _load_history(self) -> None:
        self._history_main = data_store.get_routes_history("main")
        self._history_increase = data_store.get_routes_history("increase")
        self._populate_list(self.list_main, self._history_main)
        self._populate_list(self.list_increase, self._history_increase)
        if self.list_main.count():
            self.list_main.setCurrentRow(0)
            self._last_active_list = self.list_main
        elif self.list_increase.count():
            self.list_increase.setCurrentRow(0)
            self._last_active_list = self.list_increase
        else:
            self.lbl_meta.setText("История пуста. Сначала обработайте и сохраните маршруты.")
        self._update_selection_state()

    def _populate_list(self, lst: QListWidget, history: list[dict]) -> None:
        lst.clear()
        for idx, entry in enumerate(history):
            text = _entry_text(entry)
            item = QListWidgetItem(text)
            item.setData(Qt.ItemDataRole.UserRole, ("main" if lst is self.list_main else "increase", idx))
            lst.addItem(item)

    def _get_current_entry(self) -> tuple[list[dict], int] | None:
        # Приоритет у последнего активного списка
        order = [(self.list_main, self._history_main), (self.list_increase, self._history_increase)]
        if self._last_active_list is self.list_increase:
            order.reverse()
        for lst, hist in order:
            cur = lst.currentItem()
            if cur is None or cur.isHidden():
                continue
            data = cur.data(Qt.ItemDataRole.UserRole)
            if data is None:
                continue
            _, idx = data
            if 0 <= idx < len(hist):
                return (hist, idx)
        return None

    def _apply_filter(self, text: str) -> None:
        pattern = (text or "").strip().lower()
        for lst in [self.list_main, self.list_increase]:
            first_visible = None
            for i in range(lst.count()):
                item = lst.item(i)
                visible = not pattern or pattern in (item.text() or "").lower()
                item.setHidden(not visible)
                if visible and first_visible is None:
                    first_visible = i
            cur = lst.currentItem()
            if (cur is None or cur.isHidden()) and first_visible is not None:
                lst.setCurrentRow(first_visible)
        self._update_selection_state()

    def _on_current_changed(self, current: QListWidgetItem | None, _list: QListWidget) -> None:
        self._last_active_list = _list
        self._update_selection_state()

    def _update_selection_state(self) -> None:
        res = self._get_current_entry()
        if res is None:
            self._selected = None
            self.btn_open.setEnabled(False)
            self.lbl_meta.setText("")
            return
        hist, idx = res
        entry = hist[idx]
        self._selected = entry
        created = _format_dt(entry.get("createdAt") or entry.get("timestamp"))
        modified = _format_dt(entry.get("modifiedAt") or entry.get("timestamp"))
        meta_parts = [
            f"Тип: {_type_title(entry.get('fileType', 'main'))}.",
            f"Категория: {entry.get('routeCategory') or 'ШК'}.",
            f"Маршрутов: {entry.get('count', 0)}.",
        ]
        if created:
            meta_parts.append(f"Создано: {created}.")
        if modified and modified != created:
            meta_parts.append(f"Изменено: {modified}.")
        self.lbl_meta.setText(" ".join(meta_parts))
        self.btn_open.setEnabled(True)

    def _accept_from_list(self, lst: QListWidget) -> None:
        if lst.currentItem() and self._get_current_entry():
            self._on_open_clicked()

    def _on_context_menu(self, pos) -> None:
        sender = self.sender()
        if not isinstance(sender, QListWidget):
            return
        item = sender.itemAt(pos)
        if item is None or item.isHidden():
            return
        sender.setCurrentItem(item)
        self._last_active_list = sender
        res = self._get_current_entry()
        if res is None:
            return
        menu = QMenu(self)
        act_delete = menu.addAction("Удалить")
        act = menu.exec(sender.mapToGlobal(pos))
        if act == act_delete:
            self._on_delete_selected()

    def _on_delete_selected(self) -> None:
        res = self._get_current_entry()
        if res is None:
            return
        hist, idx = res
        entry = hist[idx]
        filename = entry.get("filename")
        if not filename:
            return
        date_display = _format_dt(entry.get("modifiedAt") or entry.get("timestamp"))
        if not date_display:
            date_display = entry.get("date", "")
        reply = QMessageBox.question(
            self, "Удалить запись",
            f"Удалить сохранение от {date_display}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        if data_store.delete_routes_history_entry(filename):
            self._load_history()

    def _on_clear_history(self) -> None:
        reply = QMessageBox.question(
            self, "Очистить историю",
            "Удалить всю историю маршрутов (основной и довоз)?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        data_store.clear_last_routes()
        self._load_history()

    def _on_open_clicked(self) -> None:
        if not self._selected:
            return
        filename = self._selected.get("filename")
        if not filename:
            return
        self._loaded_entry = data_store.load_routes_history_entry(filename)
        if self._loaded_entry:
            self.accept()

    @property
    def selected_entry(self) -> dict | None:
        return deepcopy(self._loaded_entry) if self._loaded_entry else None


def pick_routes_history_entry(parent=None, file_type: str | None = None) -> dict | None:
    """Открывает диалог истории с двумя колонками. file_type игнорируется — показываются обе колонки."""
    dlg = RoutesHistoryDialog(parent)
    if dlg.exec() != QDialog.DialogCode.Accepted:
        return None
    return dlg.selected_entry
