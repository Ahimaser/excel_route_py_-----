"""
preview_general_page.py -- Предпросмотр «Общие маршруты».

Архитектура:
- QTableView + RoutesTableModel (QAbstractTableModel) вместо QTableWidget.
  Виртуальная модель не создаёт QTableWidgetItem для каждой строки --
  Qt рисует только видимые строки. 7000+ строк отображаются мгновенно.
- RenderWorker строит rows_data в фоновом потоке (без Qt-объектов).
- Данные передаются в UI-поток через сигнал finished -> слот _on_render_done
  (сигнал Qt автоматически маршалируется в UI-поток при cross-thread соединении).
- EditPanel -- боковая панель редактирования номера маршрута (без QDialog).
"""
from __future__ import annotations

import re
import os
import logging
from datetime import date, datetime, timedelta

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFrame, QTableView, QLineEdit,
    QComboBox, QHeaderView, QAbstractItemView,
    QMessageBox, QFileDialog, QProgressBar, QMenu
)
from PyQt6.QtCore import (
    Qt, pyqtSignal, QThread, QObject, QTimer,
    QAbstractTableModel, QModelIndex, QVariant
)
from PyQt6.QtGui import QFont, QColor, QBrush

from core import data_store, excel_generator
from core.xls_parser import ROUTE_SIGN

log = logging.getLogger("preview_general")

# Режимы отображения строк продуктов
_DISPLAY_FULL    = "full"
_DISPLAY_ADDR    = "addr"
_DISPLAY_PRODUCT = "product"

_UNDEFINED = "Номер маршрута не определен"

_COL_NUM  = 0
_COL_ADDR = 1
_COL_UNIT = 2
_COL_QTY  = 3
_HEADERS  = ["# маршрута", "Адрес / Продукт", "Ед. изм.", "Кол-во"]


# ─────────────────────────── Модель таблицы ───────────────────────────────

class RoutesTableModel(QAbstractTableModel):
    """Виртуальная модель -- Qt рисует только видимые строки."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._rows: list[dict] = []
        self._bold_font = QFont()
        self._bold_font.setBold(True)
        self._red_color  = QColor("#dc2626")
        self._blue_color = QColor("#2563eb")
        self._gray_bg    = QColor("#f8fafc")

    def set_rows(self, rows: list[dict]) -> None:
        self.beginResetModel()
        self._rows = rows
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()) -> int:
        return len(self._rows)

    def columnCount(self, parent=QModelIndex()) -> int:
        return 4

    def headerData(self, section: int, orientation, role=Qt.ItemDataRole.DisplayRole):
        if orientation == Qt.Orientation.Horizontal and role == Qt.ItemDataRole.DisplayRole:
            return _HEADERS[section]
        return QVariant()

    def data(self, index: QModelIndex, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return QVariant()
        row = index.row()
        col = index.column()
        if row >= len(self._rows):
            return QVariant()
        rd = self._rows[row]
        is_route = rd["type"] == "route"

        if role == Qt.ItemDataRole.DisplayRole:
            if col == _COL_NUM:
                # Читаем напрямую из route_ref чтобы видеть актуальное значение
                if rd["type"] == "route":
                    return str(rd["route_ref"].get("routeNum", ""))
                return str(rd["routeNum"])
            if col == _COL_ADDR:
                # Читаем напрямую из route_ref чтобы видеть актуальный адрес
                if rd["type"] == "route":
                    return rd["route_ref"].get("address", "")
                return rd["address"]
            if col == _COL_UNIT: return rd["unit"]
            if col == _COL_QTY:  return rd["quantity"]

        elif role == Qt.ItemDataRole.FontRole:
            if is_route:
                return self._bold_font

        elif role == Qt.ItemDataRole.ForegroundRole:
            if is_route and col == _COL_NUM:
                rnum = str(rd["route_ref"].get("routeNum", ""))
                return self._red_color if rnum == _UNDEFINED else self._blue_color

        elif role == Qt.ItemDataRole.BackgroundRole:
            if is_route:
                return self._gray_bg

        elif role == Qt.ItemDataRole.ToolTipRole:
            if is_route and col == _COL_NUM:
                return "Кликните для редактирования номера маршрута"

        elif role == Qt.ItemDataRole.SizeHintRole:
            from PyQt6.QtCore import QSize
            return QSize(-1, 36 if is_route else 30)

        return QVariant()

    def flags(self, index: QModelIndex):
        return Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable

    def get_row(self, row: int) -> dict | None:
        if 0 <= row < len(self._rows):
            return self._rows[row]
        return None

    def notify_route_changed(self, route_ref: dict) -> None:
        """Уведомляет Qt об изменении всех строк связанных с route_ref.
        Таблица перерисует только видимые строки без перестройки всей модели.
        """
        for i, rd in enumerate(self._rows):
            if rd.get("route_ref") is route_ref:
                top_left     = self.index(i, 0)
                bottom_right = self.index(i, self.columnCount() - 1)
                self.dataChanged.emit(top_left, bottom_right,
                                      [Qt.ItemDataRole.DisplayRole,
                                       Qt.ItemDataRole.ForegroundRole])


# ─────────────────────────── Worker рендера ───────────────────────────────

class RenderWorker(QObject):
    """Строит rows_data в фоновом потоке -- не блокирует UI."""
    finished = pyqtSignal(list, int, int)  # (rows_data, visible_count, no_num_count)

    def __init__(self, routes: list, prod_settings: dict,
                 search_lower: str, filter_prod: str, display_mode: str,
                 sort_asc: bool = False):
        super().__init__()
        self.routes        = routes
        self.prod_settings = prod_settings
        self.search_lower  = search_lower
        self.filter_prod   = filter_prod
        self.display_mode  = display_mode
        self.sort_asc      = sort_asc

    def run(self) -> None:
        try:
            result = _build_rows(
                self.routes, self.prod_settings,
                self.search_lower, self.filter_prod, self.display_mode,
                self.sort_asc
            )
            self.finished.emit(*result)
        except Exception as exc:
            log.exception("RenderWorker error: %s", exc)
            self.finished.emit([], 0, 0)


def _build_rows(routes: list, prod_settings: dict,
                search_lower: str, filter_prod: str,
                display_mode: str, sort_asc: bool = False) -> tuple[list, int, int]:
    """Чистая функция -- строит rows_data без Qt-объектов."""

    def _sort_key(r: dict):
        num = r.get("routeNum", "")
        if num == _UNDEFINED or not str(num).strip():
            # Неопределённые всегда в начало (независимо от направления)
            return (0, 0)
        try:
            n = int(str(num).strip())
            return (1, n if sort_asc else -n)
        except ValueError:
            return (1, 0)

    def _fmt_qty(prod: dict) -> str:
        qty = prod.get("quantity")
        if qty is None:
            return ""
        ps = prod_settings.get(prod["name"], {})
        if ps.get("showPcs") and ps.get("pcsPerUnit", 0) > 0:
            pcs = excel_generator.calc_pcs(
                qty, ps["pcsPerUnit"], ps.get("roundUp", True)
            )
            return f"{qty} / {pcs} шт"
        return str(qty)

    sorted_routes = sorted(routes, key=_sort_key)
    rows_data: list[dict] = []
    visible_count = 0
    no_num_count  = 0

    for r in sorted_routes:
        if r.get("excluded"):
            continue
        if search_lower:
            addr = r.get("address", "").lower()
            num  = str(r.get("routeNum", "")).lower()
            if search_lower not in addr and search_lower not in num:
                continue
        if filter_prod:
            if not any(p["name"] == filter_prod for p in r.get("products", [])):
                continue

        visible_count += 1
        rnum = r.get("routeNum", "")
        if rnum == _UNDEFINED or not str(rnum).strip():
            no_num_count += 1

        rows_data.append({
            "type":      "route",
            "routeNum":  rnum,
            "address":   r.get("address", ""),
            "unit":      "",
            "quantity":  "",
            "route_ref": r,
        })

        if display_mode == _DISPLAY_ADDR:
            pass
        elif display_mode == _DISPLAY_PRODUCT and filter_prod:
            for p in r.get("products", []):
                if p["name"] == filter_prod:
                    rows_data.append({
                        "type":      "product",
                        "routeNum":  "",
                        "address":   f"  {p['name']}",
                        "unit":      p.get("unit", ""),
                        "quantity":  _fmt_qty(p),
                        "route_ref": r,
                    })
        else:
            for p in r.get("products", []):
                rows_data.append({
                    "type":      "product",
                    "routeNum":  "",
                    "address":   f"  {p['name']}",
                    "unit":      p.get("unit", ""),
                    "quantity":  _fmt_qty(p),
                    "route_ref": r,
                })

    return rows_data, visible_count, no_num_count


# ─────────────────────────── Worker генерации ─────────────────────────────

class GenerateWorker(QObject):
    finished = pyqtSignal(str)
    error    = pyqtSignal(str)

    def __init__(self, routes: list, file_type: str, save_path: str, prod_map: dict):
        super().__init__()
        # Делаем глубокую копию маршрутов чтобы избежать конкурентного доступа
        import copy
        self.routes    = copy.deepcopy(routes)
        self.file_type = file_type
        self.save_path = save_path
        self.prod_map  = prod_map

    def run(self) -> None:
        try:
            path = excel_generator.generate_general_routes(
                self.routes, self.file_type, self.save_path, self.prod_map
            )
            self.finished.emit(path)
        except Exception as exc:
            self.error.emit(str(exc))


# ─────────────────────────── Боковая панель редактирования ────────────────

class EditPanel(QFrame):
    """Боковая панель для редактирования номера маршрута.

    Не использует QDialog -- полностью встроена в layout страницы.
    Сигнал saved(route_ref, new_num) испускается при успешном сохранении.
    """

    saved  = pyqtSignal(object, str)
    closed = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("editPanel")
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setMinimumWidth(260)
        self.setMaximumWidth(320)
        self.setStyleSheet(
            "#editPanel {"
            "  background: #f8fafc;"
            "  border-left: 2px solid #e2e8f0;"
            "}"
        )
        self._route_ref: dict | None = None

        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 16, 16, 16)
        lay.setSpacing(12)

        title_row = QHBoxLayout()
        lbl_title = QLabel("Редактировать номер")
        lbl_title.setStyleSheet("font-weight: bold; font-size: 13px; color: #1e293b;")
        title_row.addWidget(lbl_title)
        title_row.addStretch()
        btn_close = QPushButton("x")
        btn_close.setFixedSize(24, 24)
        btn_close.setStyleSheet(
            "QPushButton { background: transparent; border: none; color: #94a3b8;"
            "  font-size: 14px; font-weight: bold; }"
            "QPushButton:hover { color: #dc2626; }"
        )
        btn_close.clicked.connect(self._on_close)
        title_row.addWidget(btn_close)
        lay.addLayout(title_row)

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color: #e2e8f0;")
        lay.addWidget(sep)

        lbl_addr_caption = QLabel("Адрес:")
        lbl_addr_caption.setStyleSheet("color: #64748b; font-size: 11px;")
        lay.addWidget(lbl_addr_caption)

        self.lbl_address = QLabel("")
        self.lbl_address.setWordWrap(True)
        self.lbl_address.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        self.lbl_address.setStyleSheet(
            "background: #fff; border: 1px solid #e2e8f0; border-radius: 6px;"
            "padding: 8px; font-size: 12px; color: #1e293b;"
        )
        lay.addWidget(self.lbl_address)

        lbl_cur_caption = QLabel("Текущий номер:")
        lbl_cur_caption.setStyleSheet("color: #64748b; font-size: 11px;")
        lay.addWidget(lbl_cur_caption)

        self.lbl_current = QLabel("")
        self.lbl_current.setStyleSheet(
            "font-size: 22px; font-weight: bold; color: #2563eb; padding: 4px 0px;"
        )
        lay.addWidget(self.lbl_current)

        lbl_new_caption = QLabel("Новый номер:")
        lbl_new_caption.setStyleSheet("color: #64748b; font-size: 11px;")
        lay.addWidget(lbl_new_caption)

        self.le_new_num = QLineEdit()
        self.le_new_num.setPlaceholderText("Введите число...")
        self._style_normal = (
            "QLineEdit { border: 2px solid #2563eb; border-radius: 6px;"
            "  padding: 8px 10px; font-size: 16px; background: #fff; }"
            "QLineEdit:focus { border-color: #1d4ed8; }"
        )
        self._style_error = (
            "QLineEdit { border: 2px solid #dc2626; border-radius: 6px;"
            "  padding: 8px 10px; font-size: 16px; background: #fff; }"
        )
        self.le_new_num.setStyleSheet(self._style_normal)
        self.le_new_num.returnPressed.connect(self._on_save)
        lay.addWidget(self.le_new_num)

        btn_row = QHBoxLayout()
        self.btn_save = QPushButton("Сохранить")
        self.btn_save.setObjectName("btnPrimary")
        self.btn_save.setFixedHeight(36)
        self.btn_save.clicked.connect(self._on_save)
        btn_row.addWidget(self.btn_save)

        btn_cancel = QPushButton("Отмена")
        btn_cancel.setObjectName("btnSecondary")
        btn_cancel.setFixedHeight(36)
        btn_cancel.clicked.connect(self._on_close)
        btn_row.addWidget(btn_cancel)
        lay.addLayout(btn_row)

        lay.addStretch()

    def load(self, route_ref: dict) -> None:
        self._route_ref = route_ref
        address = route_ref.get("address", "")
        rnum    = str(route_ref.get("routeNum", ""))
        self.lbl_address.setText(address)
        self.lbl_current.setText(rnum if rnum != _UNDEFINED else "—")
        self.le_new_num.setText("" if rnum == _UNDEFINED else rnum)
        self.le_new_num.setStyleSheet(self._style_normal)
        self.le_new_num.selectAll()
        self.le_new_num.setFocus()
        self.setVisible(True)

    def update_current(self, new_num: str) -> None:
        """Обновляет отображение текущего номера после сохранения."""
        self.lbl_current.setText(new_num)

    def clear(self) -> None:
        self._route_ref = None
        self.lbl_address.setText("")
        self.lbl_current.setText("")
        self.le_new_num.clear()
        self.setVisible(False)

    def _on_save(self) -> None:
        if self._route_ref is None:
            return
        new_val = self.le_new_num.text().strip()
        if not new_val:
            self._on_close()
            return
        if not re.match(r"^\d+$", new_val):
            self.le_new_num.setStyleSheet(self._style_error)
            self.le_new_num.setFocus()
            self.le_new_num.selectAll()
            return
        self.le_new_num.setStyleSheet(self._style_normal)
        route = self._route_ref
        self.saved.emit(route, new_val)
        # Обновляем отображение текущего номера сразу после сохранения
        self.lbl_current.setText(new_val)

    def _on_close(self) -> None:
        self.clear()
        self.closed.emit()


# ─────────────────────────── Страница ─────────────────────────────────────

class PreviewGeneralPage(QWidget):
    """Предпросмотр и генерация файла «Общие маршруты»."""

    go_back         = pyqtSignal()
    go_dept_preview = pyqtSignal()
    go_settings     = pyqtSignal()   # Переход к настройкам Шт

    def __init__(self, app_state: dict):
        super().__init__()
        self.app_state       = app_state
        self._display_mode   = _DISPLAY_FULL
        self._filter_product = ""
        self._search_text    = ""
        self._rendering      = False
        self._render_pending = False
        self._sort_asc       = False  # False = убывание (по умолчанию)

        self._render_thread: QThread | None = None
        self._render_worker: RenderWorker | None = None

        # Debounce для поиска
        self._search_timer = QTimer(self)
        self._search_timer.setSingleShot(True)
        self._search_timer.timeout.connect(self._render_table)

        self._build_ui()

    # ─────────────────────────── Построение UI ────────────────────────────

    def _build_ui(self) -> None:
        root_lay = QVBoxLayout(self)
        root_lay.setContentsMargins(24, 16, 24, 16)
        root_lay.setSpacing(12)

        # Заголовок
        h_row = QHBoxLayout()
        btn_back = QPushButton("< Назад")
        btn_back.setObjectName("btnBack")
        btn_back.clicked.connect(self.go_back.emit)
        h_row.addWidget(btn_back)

        self.lbl_title = QLabel("Общие маршруты")
        self.lbl_title.setObjectName("sectionTitle")
        h_row.addWidget(self.lbl_title)
        h_row.addStretch()

        self.lbl_count = QLabel("")
        self.lbl_count.setObjectName("badge")
        h_row.addWidget(self.lbl_count)

        self.lbl_no_num = QLabel("")
        self.lbl_no_num.setObjectName("badgeRed")
        self.lbl_no_num.setVisible(False)
        h_row.addWidget(self.lbl_no_num)

        btn_settings = QPushButton("⚙ Настройки Шт")
        btn_settings.setObjectName("btnSecondary")
        btn_settings.setToolTip("Открыть настройки отображения в штуках")
        btn_settings.clicked.connect(self.go_settings.emit)
        h_row.addWidget(btn_settings)

        root_lay.addLayout(h_row)

        # Панель фильтров
        filter_card = QFrame()
        filter_card.setObjectName("card")
        filter_lay = QHBoxLayout(filter_card)
        filter_lay.setContentsMargins(12, 10, 12, 10)
        filter_lay.setSpacing(12)

        self.le_search = QLineEdit()
        self.le_search.setPlaceholderText("Поиск по адресу или номеру маршрута...")
        self.le_search.setMinimumWidth(240)
        self.le_search.textChanged.connect(self._on_search_changed)
        filter_lay.addWidget(self.le_search)

        filter_lay.addWidget(QLabel("Продукт:"))
        self.combo_product = QComboBox()
        self.combo_product.setMinimumWidth(180)
        self.combo_product.addItem("Все продукты", "")
        self.combo_product.currentIndexChanged.connect(self._on_product_filter_changed)
        filter_lay.addWidget(self.combo_product)

        filter_lay.addWidget(QLabel("Фильтр:"))
        self.combo_display = QComboBox()
        self.combo_display.setMinimumWidth(160)
        self.combo_display.addItem("Полностью",      _DISPLAY_FULL)
        self.combo_display.addItem("Только адреса",  _DISPLAY_ADDR)
        self.combo_display.addItem("Только продукт", _DISPLAY_PRODUCT)
        self.combo_display.setToolTip(
            "Полностью -- маршруты и все строки продуктов\n"
            "Только адреса -- скрыть строки продуктов\n"
            "Только продукт -- показать строки только выбранного продукта"
        )
        self.combo_display.currentIndexChanged.connect(self._on_display_changed)
        filter_lay.addWidget(self.combo_display)

        # Кнопка сортировки
        self.btn_sort = QPushButton("↓ По убыванию")
        self.btn_sort.setObjectName("btnSecondary")
        self.btn_sort.setToolTip("Сортировка по номеру маршрута")
        self.btn_sort.clicked.connect(self._on_sort_toggle)
        filter_lay.addWidget(self.btn_sort)

        filter_lay.addStretch()

        btn_reset = QPushButton("Сбросить")
        btn_reset.setObjectName("btnSecondary")
        btn_reset.clicked.connect(self._reset_filters)
        filter_lay.addWidget(btn_reset)

        root_lay.addWidget(filter_card)

        # Таблица + боковая панель
        content_row = QHBoxLayout()
        content_row.setSpacing(0)

        # QTableView с виртуальной моделью
        self._model = RoutesTableModel(self)
        self.table = QTableView()
        self.table.setModel(self._model)
        hdr = self.table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.clicked.connect(self._on_cell_clicked)
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._on_context_menu)
        content_row.addWidget(self.table, stretch=1)

        # Боковая панель
        self.edit_panel = EditPanel(self)
        self.edit_panel.setVisible(False)
        self.edit_panel.saved.connect(self._on_route_num_saved)
        self.edit_panel.closed.connect(self._on_panel_closed)
        content_row.addWidget(self.edit_panel)

        root_lay.addLayout(content_row, stretch=1)

        # Нижняя панель
        bottom_row = QHBoxLayout()
        self.lbl_excluded = QLabel("")
        self.lbl_excluded.setStyleSheet("color: #64748b; font-size: 12px;")
        bottom_row.addWidget(self.lbl_excluded)
        bottom_row.addStretch()

        self.btn_generate = QPushButton("Создать файл «Общие маршруты»")
        self.btn_generate.setObjectName("btnPrimary")
        self.btn_generate.setFixedHeight(40)
        self.btn_generate.clicked.connect(self._on_generate)
        bottom_row.addWidget(self.btn_generate)

        self.btn_labels = QPushButton("Этикетки из шаблонов (XLS)")
        self.btn_labels.setObjectName("btnSecondary")
        self.btn_labels.setFixedHeight(40)
        self.btn_labels.setToolTip("Создать этикетки в папку «Этикетки на ДД.ММ.ГГГГ» (завтра)")
        self.btn_labels.clicked.connect(self._on_labels_from_templates)
        bottom_row.addWidget(self.btn_labels)

        root_lay.addLayout(bottom_row)

        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.setVisible(False)
        root_lay.addWidget(self.progress)

    # ─────────────────────────── Обновление ───────────────────────────────

    def refresh(self) -> None:
        log.debug("refresh called")
        file_type = self.app_state.get("fileType", "main")
        self.lbl_title.setText(
            "Общие маршруты — "
            + ("Основной" if file_type == "main" else "Увеличение (Довоз)")
        )
        self.combo_product.blockSignals(True)
        self.combo_product.clear()
        self.combo_product.addItem("Все продукты", "")
        for p in sorted(
            self.app_state.get("uniqueProducts", []), key=lambda x: x["name"]
        ):
            self.combo_product.addItem(p["name"], p["name"])
        self.combo_product.blockSignals(False)
        self.edit_panel.clear()
        self._render_table()

    # ─────────────────────────── Рендер ───────────────────────────────────

    def _render_table(self) -> None:
        log.debug("_render_table called, rendering=%s", self._rendering)
        if self._rendering:
            self._render_pending = True
            return
        self._start_render_worker()

    def _start_render_worker(self) -> None:
        self._rendering = True
        self._render_pending = False

        routes        = list(self.app_state.get("filteredRoutes", []))
        prod_settings = data_store.get_products_map()
        search_lower  = self._search_text.lower()
        filter_prod   = self._filter_product
        display_mode  = self._display_mode
        sort_asc      = self._sort_asc

        log.debug("RenderWorker start: %d routes, mode=%s, sort=%s",
                  len(routes), display_mode, 'asc' if sort_asc else 'desc')

        self._render_thread = QThread(self)
        self._render_worker = RenderWorker(
            routes, prod_settings, search_lower, filter_prod, display_mode, sort_asc
        )
        self._render_worker.moveToThread(self._render_thread)
        self._render_thread.started.connect(self._render_worker.run)
        # Qt автоматически маршалирует сигнал finished в UI-поток
        self._render_worker.finished.connect(self._on_render_done)
        self._render_worker.finished.connect(self._render_thread.quit)
        self._render_thread.finished.connect(self._render_worker.deleteLater)
        self._render_thread.finished.connect(self._render_thread.deleteLater)
        self._render_thread.start()

    def _on_render_done(self, rows_data: list, visible_count: int,
                        no_num_count: int) -> None:
        """Вызывается в UI-потоке после завершения RenderWorker."""
        log.debug("_on_render_done: %d rows", len(rows_data))
        self._rendering = False

        # Сохраняем текущий route_ref из edit_panel перед обновлением модели
        current_route_ref = self.edit_panel._route_ref

        # Обновляем виртуальную модель -- мгновенно, без создания виджетов
        self._model.set_rows(rows_data)

        # Если панель открыта, обновляем адрес и прокручиваем к изменённому маршруту
        if current_route_ref is not None and self.edit_panel.isVisible():
            self.edit_panel.lbl_address.setText(current_route_ref.get("address", ""))
            # Находим строку изменённого маршрута в новых данных и прокручиваем к ней
            for i, rd in enumerate(rows_data):
                if rd.get("route_ref") is current_route_ref:
                    idx = self._model.index(i, 0)
                    self.table.scrollTo(idx)
                    self.table.setCurrentIndex(idx)
                    break

        all_routes     = self.app_state.get("filteredRoutes", [])
        excluded_count = sum(1 for r in all_routes if r.get("excluded"))
        self.lbl_count.setText(f"{visible_count} маршрутов")
        self.lbl_excluded.setText(
            f"Исключено: {excluded_count}" if excluded_count else ""
        )
        self._update_no_num_label(no_num_count)

        log.debug("_on_render_done done")

        if self._render_pending:
            log.debug("render_pending=True, запускаем ещё раз")
            self._start_render_worker()

    # ─────────────────────────── Фильтры ──────────────────────────────────

    def _on_search_changed(self, text: str) -> None:
        self._search_text = text
        self._search_timer.start(200)

    def _on_product_filter_changed(self) -> None:
        self._filter_product = self.combo_product.currentData() or ""
        self._render_table()

    def _on_display_changed(self) -> None:
        self._display_mode = self.combo_display.currentData() or _DISPLAY_FULL
        self._render_table()

    def _on_sort_toggle(self) -> None:
        self._sort_asc = not self._sort_asc
        if self._sort_asc:
            self.btn_sort.setText("↑ По возрастанию")
        else:
            self.btn_sort.setText("↓ По убыванию")
        # Сохраняем порядок сортировки в app_state для использования в dept-файлах
        self.app_state["sortAsc"] = self._sort_asc
        self._render_table()

    def _reset_filters(self) -> None:
        self.le_search.blockSignals(True)
        self.le_search.clear()
        self.le_search.blockSignals(False)
        self.combo_product.blockSignals(True)
        self.combo_product.setCurrentIndex(0)
        self.combo_product.blockSignals(False)
        self.combo_display.blockSignals(True)
        self.combo_display.setCurrentIndex(0)
        self.combo_display.blockSignals(False)
        self._search_text    = ""
        self._filter_product = ""
        self._display_mode   = _DISPLAY_FULL
        self._sort_asc       = False
        self.btn_sort.setText("↓ По убыванию")
        self.app_state["sortAsc"] = False
        self.edit_panel.clear()
        self._render_table()

    # ─────────────────────────── Клик по строке ───────────────────────────

    def _on_cell_clicked(self, index: QModelIndex) -> None:
        rd = self._model.get_row(index.row())
        if rd is None:
            return
        if rd["type"] != "route":
            self.edit_panel.clear()
            return
        self.edit_panel.load(rd["route_ref"])

    def _on_route_num_saved(self, route: dict, new_val: str) -> None:
        old_num = str(route.get("routeNum", ""))
        if new_val == old_num:
            log.debug("route_num_saved: значение не изменилось ('%s')", old_num)
            return

         # Обновляем оригинальный словарь в filteredRoutes
        # route_ref в rows_data — это прямая ссылка на этот объект
        route["routeNum"] = new_val
        log.debug("route_num_saved: '%s' -> '%s', id=%d", old_num, new_val, id(route))
        # Проверяем что объект действительно в filteredRoutes
        all_routes = self.app_state.get("filteredRoutes", [])
        found_idx = next((i for i, r in enumerate(all_routes) if r is route), -1)
        if found_idx == -1:
            # route_ref указывает на устаревший объект (не из filteredRoutes).
            # Ищем маршрут по адресу и обновляем правильный объект.
            route_address = route.get("address", "")
            fallback_idx = next(
                (i for i, r in enumerate(all_routes)
                 if r.get("address", "") == route_address),
                -1
            )
            if fallback_idx != -1:
                real_route = all_routes[fallback_idx]
                real_route["routeNum"] = new_val
                # Перенаправляем route на реальный объект для дальнейших операций
                route = real_route
                log.debug("route_num_saved: fallback по address, обновлён filteredRoutes[%d]",
                          fallback_idx)
            else:
                log.warning("route_num_saved: маршрут НЕ найден ни по id ни по address — "
                            "изменение потеряется")
        else:
            log.debug("  маршрут найден в filteredRoutes[%d], routeNum=%s",
                      found_idx, all_routes[found_idx].get("routeNum"))

        # Обновляем номер в адресной строке
        old_address = route.get("address", "")
        idx = old_address.find(ROUTE_SIGN)
        if idx != -1:
            tail = old_address[idx + 1:].strip()
            m = re.match(r"(\d+)(.*)", tail)
            if m:
                route["address"] = old_address[:idx + 1] + new_val + m.group(2)
            else:
                route["address"] = old_address[:idx + 1] + new_val

        # Обновляем адрес в панели (номер уже обновлён в EditPanel._on_save)
        self.edit_panel.lbl_address.setText(route.get("address", ""))

        # Немедленно уведомляем модель об изменении строк этого маршрута
        # (без полного пересчёта rows_data)
        self._model.notify_route_changed(route)

        # Перезапускаем рендер — пересортировка по новому номеру
        self._render_table()

        # Обновляем счётчик неопределённых номеров
        no_num_count = sum(
            1 for r in all_routes
            if not r.get("excluded")
            and (r.get("routeNum") == _UNDEFINED or not str(r.get("routeNum", "")).strip())
        )
        self._update_no_num_label(no_num_count)
        log.debug("route_num_saved done, no_num_count=%d", no_num_count)

    def _update_no_num_label(self, no_num_count: int) -> None:
        """Обновляет метку неопределённых номеров маршрутов."""
        if no_num_count > 0:
            self.lbl_no_num.setText(f"Номер не определён: {no_num_count}")
            self.lbl_no_num.setObjectName("badgeRed")
            self.lbl_no_num.setVisible(True)
        else:
            self.lbl_no_num.setText("Номера маршрутов определены")
            self.lbl_no_num.setObjectName("badgeGreen")
            self.lbl_no_num.setVisible(True)
        # Принудительно обновляем стиль (objectName изменился)
        self.lbl_no_num.style().unpolish(self.lbl_no_num)
        self.lbl_no_num.style().polish(self.lbl_no_num)

    def _on_panel_closed(self) -> None:
        self.table.clearSelection()

    # ─────────────────────────── Контекстное меню ─────────────────────────

    def _on_context_menu(self, pos) -> None:
        index = self.table.indexAt(pos)
        if not index.isValid():
            return
        rd = self._model.get_row(index.row())
        if rd is None or rd["type"] != "route":
            return

        menu  = QMenu(self)
        route = rd["route_ref"]

        if route.get("excluded"):
            act_toggle = menu.addAction("Включить маршрут")
        else:
            act_toggle = menu.addAction("Исключить маршрут")

        act_edit = menu.addAction("Изменить номер маршрута...")

        action = menu.exec(self.table.viewport().mapToGlobal(pos))

        if action == act_toggle:
            route["excluded"] = not route.get("excluded", False)
            self._render_table()
        elif action == act_edit:
            self.edit_panel.load(route)

    # ─────────────────────────── Генерация ────────────────────────────────────

    def _on_generate(self) -> None:
        routes        = self.app_state.get("filteredRoutes", [])
        active_routes = [r for r in routes if not r.get("excluded")]

        if not active_routes:
            QMessageBox.warning(self, "Нет маршрутов",
                                "Нет маршрутов для создания файла.")
            return

        # Проверяем маршруты с неопределённым или пустым номером
        undefined = [
            r for r in active_routes
            if r.get("routeNum") == _UNDEFINED
            or not str(r.get("routeNum", "")).strip()
        ]
        if undefined:
            msg = f"Найдено {len(undefined)} маршрут(ов) с неопределённым номером:\n\n"
            for r in undefined[:5]:
                msg += f"  - {r.get('address', '')[:60]}\n"
            if len(undefined) > 5:
                msg += f"  ... и ещё {len(undefined) - 5}\n"
            msg += "\nПродолжить генерацию файла?"
            reply = QMessageBox.question(
                self, "Неопределённые номера маршрутов", msg,
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply != QMessageBox.StandardButton.Yes:
                return

        tomorrow     = datetime.now() + timedelta(days=1)
        date_str     = tomorrow.strftime("%d.%m.%Y")
        file_type    = self.app_state.get("fileType", "main")
        type_label   = "ОСН" if file_type == "main" else "УВ"
        default_name = f"Маршруты общие {date_str} {type_label}.xls"

        save_dir  = self.app_state.get("saveDir") or data_store.get_desktop_path()
        save_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить файл",
            os.path.join(save_dir, default_name),
            "Excel 97-2003 (*.xls)"
        )
        if not save_path:
            return
        if not save_path.lower().endswith(".xls"):
            save_path += ".xls"

        if os.path.exists(save_path):
            reply = QMessageBox.question(
                self, "Файл существует",
                f"Файл уже существует:\n{save_path}\n\nПерезаписать?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply != QMessageBox.StandardButton.Yes:
                return

        prod_map = data_store.get_products_map()
        self.btn_generate.setEnabled(False)
        self.progress.setVisible(True)

        self._gen_thread = QThread(self)
        self._gen_worker = GenerateWorker(active_routes, file_type, save_path, prod_map)
        self._gen_worker.moveToThread(self._gen_thread)
        self._gen_thread.started.connect(self._gen_worker.run)
        self._gen_worker.finished.connect(self._on_gen_done)
        self._gen_worker.error.connect(self._on_gen_error)
        self._gen_worker.finished.connect(self._gen_thread.quit)
        self._gen_worker.error.connect(self._gen_thread.quit)
        self._gen_thread.finished.connect(self._gen_worker.deleteLater)
        self._gen_thread.start()

    def _on_gen_done(self, path: str) -> None:
        self.progress.setVisible(False)
        self.btn_generate.setEnabled(True)
        file_type = self.app_state.get("fileType", "main")
        data_store.save_last_routes(
            file_type,
            self.app_state.get("routes", []),
            self.app_state.get("uniqueProducts", []),
            self.app_state.get("filteredRoutes", []),
            route_category=self.app_state.get("routeCategory"),
        )
        reply = QMessageBox.information(
            self, "Готово",
            f"Файл успешно создан:\n{path}\n\n"
            "Перейти к предпросмотру файлов по отделам?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.go_dept_preview.emit()

    def _on_gen_error(self, msg: str) -> None:
        self.progress.setVisible(False)
        self.btn_generate.setEnabled(True)
        QMessageBox.critical(self, "Ошибка",
                             f"Ошибка при создании файла:\n{msg}")

    def _on_labels_from_templates(self) -> None:
        routes = self.app_state.get("filteredRoutes", [])
        active = [r for r in routes if not r.get("excluded")]
        if not active:
            QMessageBox.warning(self, "Нет данных", "Нет маршрутов для этикеток.")
            return
        products_ref = data_store.get_ref("products") or []
        if not any(p.get("labelTemplatePath") for p in products_ref):
            QMessageBox.information(
                self, "Нет шаблонов",
                "Откройте «Настройки этикеток» (страница Этикетки) и выберите шаблон XLS для продуктов."
            )
            return
        base_dir = self.app_state.get("saveDir") or data_store.get_desktop_path()
        tomorrow = date.today() + timedelta(days=1)
        folder_name = f"Этикетки на {tomorrow:%d.%m.%Y}"
        out_dir = os.path.join(base_dir, folder_name)
        os.makedirs(out_dir, exist_ok=True)
        file_type = self.app_state.get("fileType", "main")
        departments_ref = data_store.get_ref("departments") or []
        try:
            created = excel_generator.generate_labels_from_templates(
                routes, out_dir, file_type, products_ref, departments_ref
            )
            if created:
                QMessageBox.information(self, "Готово", f"Создано файлов: {len(created)}\n\n{out_dir}")
            else:
                QMessageBox.information(self, "Нет файлов", "Нет этикеток для создания.")
        except Exception as e:
            log.exception("labels")
            QMessageBox.critical(self, "Ошибка", str(e))
