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
    QFrame, QTableView, QLineEdit, QApplication,
    QComboBox, QHeaderView, QAbstractItemView,
    QMessageBox, QFileDialog, QProgressBar, QMenu, QScrollArea,
    QStyledItemDelegate,
)
from PyQt6.QtCore import (
    Qt, pyqtSignal, QThread, QObject, QTimer, QEvent,
    QAbstractTableModel, QModelIndex, QVariant
)
from PyQt6.QtGui import QFont, QColor, QBrush, QWheelEvent, QShortcut, QKeySequence

from core import data_store, excel_generator
from ui.excel_safe_open import open_excel_file_safely
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


# ─────────────────────────── Делегат (заливка из модели) ──────────────────

class RoutesTableDelegate(QStyledItemDelegate):
    """Отрисовывает фон из BackgroundRole модели (стили Qt иначе переопределяют)."""

    def paint(self, painter, option, index):
        bg = index.data(Qt.ItemDataRole.BackgroundRole)
        if bg is not None:
            try:
                brush = QBrush(bg) if not isinstance(bg, QBrush) else bg
                if brush.style() != Qt.BrushStyle.NoBrush:
                    painter.save()
                    painter.fillRect(option.rect, brush)
                    painter.restore()
            except (TypeError, ValueError):
                pass
        super().paint(painter, option, index)


# ─────────────────────────── Модель таблицы ───────────────────────────────

class RoutesTableModel(QAbstractTableModel):
    """Виртуальная модель -- Qt рисует только видимые строки."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._rows: list[dict] = []
        self._font_size = 11
        self._bold_font = QFont()
        self._bold_font.setBold(True)
        self._bold_font.setPointSize(self._font_size)
        self._red_color  = QColor("#dc2626")
        self._blue_color = QColor("#2563eb")
        self._gray_bg    = QColor("#f8fafc")
        self._red_bg     = QColor("#FEE2E2")   # заливка для «не определен»
        self._blue_bg    = QColor("#DBEAFE")   # заливка для «определен»

    def set_font_size(self, point_size: int) -> None:
        """Меняет размер шрифта (для Ctrl+колесо мыши)."""
        self._font_size = max(8, min(24, point_size))
        self._bold_font.setPointSize(self._font_size)

    def get_font_size(self) -> int:
        return self._font_size

    def emit_data_changed(self) -> None:
        """Обновляет отображение после смены шрифта."""
        if self._rows:
            top = self.index(0, 0)
            bottom = self.index(len(self._rows) - 1, 3)
            self.dataChanged.emit(top, bottom, [Qt.ItemDataRole.FontRole])

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
                if col == _COL_NUM:
                    rnum = str(rd["route_ref"].get("routeNum", ""))
                    return self._red_bg if rnum == _UNDEFINED else self._blue_bg
                return self._gray_bg

        elif role == Qt.ItemDataRole.ToolTipRole:
            if is_route and col == _COL_NUM:
                return "Кликните для редактирования номера маршрута"

        elif role == Qt.ItemDataRole.SizeHintRole:
            from PyQt6.QtCore import QSize
            base = max(32, self._font_size + 16)
            return QSize(-1, base + 10 if is_route else base + 4)

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
                                       Qt.ItemDataRole.ForegroundRole,
                                       Qt.ItemDataRole.BackgroundRole])


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

    def _fmt_qty(prod: dict, route: dict) -> str:
        qty = prod.get("quantity")
        if qty is None:
            return ""
        ps = prod_settings.get(prod["name"], {})
        mult = float(ps.get("quantityMultiplier", 1.0) or 1.0)
        try:
            display_qty = float(qty) * mult
        except (ValueError, TypeError):
            display_qty = qty
        if ps.get("showPcs") and ps.get("pcsPerUnit", 0) > 0:
            route_cat = route.get("routeCategory") or "ШК"
            addr = route.get("address", "")
            force_round_up = excel_generator.is_always_round_up_institution(addr)
            pcu = float(ps["pcsPerUnit"])
            if force_round_up:
                pct = excel_generator.get_institution_round_percent(ps.get("deptKey"))
                round_tail = pcu * (pct / 100.0)
                pcs = excel_generator.calc_pcs_tail(display_qty, pcu, round_tail)
            else:
                round_tail = (
                    ps.get("roundTailFromСД") if route_cat == "СД" else ps.get("roundTailFromШК")
                )
                if round_tail is not None:
                    pcs = excel_generator.calc_pcs_tail(
                        display_qty, pcu, float(round_tail)
                    )
                else:
                    round_up = (
                        ps.get("roundUpСД") if "roundUpСД" in ps else ps.get("roundUp", True)
                        if route_cat == "СД"
                        else ps.get("roundUpШК") if "roundUpШК" in ps else ps.get("roundUp", True)
                    )
                    pcs = excel_generator.calc_pcs(display_qty, pcu, bool(round_up))
            tail = prod.get("pcsTail")
            if tail is not None and (prod.get("unit") or ""):
                try:
                    tval = float(tail)
                    if tval > 1e-9:
                        ttxt = str(int(tval)) if abs(tval - round(tval)) < 1e-9 else f"{tval:.3f}".rstrip("0").rstrip(".")
                        return f"{display_qty} / {pcs} шт + {ttxt} {prod.get('unit', '')}"
                except (TypeError, ValueError):
                    pass
            return f"{display_qty} / {pcs} шт"
        return str(display_qty)

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
                        "quantity":  _fmt_qty(p, r),
                        "route_ref": r,
                    })
        else:
            for p in r.get("products", []):
                rows_data.append({
                    "type":      "product",
                    "routeNum":  "",
                    "address":   f"  {p['name']}",
                    "unit":      p.get("unit", ""),
                    "quantity":  _fmt_qty(p, r),
                    "route_ref": r,
                })

    return rows_data, visible_count, no_num_count


# ─────────────────────────── Worker генерации ─────────────────────────────

class GenerateWorker(QObject):
    finished = pyqtSignal(str)
    error    = pyqtSignal(str)

    def __init__(self, routes: list, file_type: str, save_path: str, prod_map: dict,
                 sort_asc: bool = True, date_str: str | None = None):
        super().__init__()
        # Делаем глубокую копию маршрутов чтобы избежать конкурентного доступа
        import copy
        self.routes    = copy.deepcopy(routes)
        self.file_type = file_type
        self.save_path = save_path
        self.prod_map  = prod_map
        self.sort_asc  = sort_asc
        self.date_str  = date_str

    def run(self) -> None:
        try:
            path = excel_generator.generate_general_routes(
                self.routes, self.file_type, self.save_path, self.prod_map, self.sort_asc,
                date_str=self.date_str
            )
            self.finished.emit(path)
        except Exception as exc:
            self.error.emit(str(exc))


class TemplateLabelsWorker(QObject):
    finished = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(self, routes: list, out_dir: str, file_type: str, products_ref: list, departments_ref: list):
        super().__init__()
        import copy
        self.routes = copy.deepcopy(routes)
        self.out_dir = out_dir
        self.file_type = file_type
        self.products_ref = products_ref
        self.departments_ref = departments_ref

    def run(self) -> None:
        try:
            created = excel_generator.generate_labels_from_templates(
                self.routes, self.out_dir, self.file_type, self.products_ref, self.departments_ref
            )
            self.finished.emit(created)
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
        self.setMinimumWidth(220)
        self.setMaximumWidth(320)
        self._route_ref: dict | None = None

        lay = QVBoxLayout(self)
        lay.setContentsMargins(12, 12, 12, 12)
        lay.setSpacing(8)

        title_row = QHBoxLayout()
        lbl_title = QLabel("Редактировать номер")
        lbl_title.setObjectName("panelTitle")
        title_row.addWidget(lbl_title)
        title_row.addStretch()
        btn_close = QPushButton("×")
        btn_close.setObjectName("btnPanelClose")
        btn_close.setFixedSize(24, 24)
        btn_close.clicked.connect(self._on_close)
        title_row.addWidget(btn_close)
        lay.addLayout(title_row)

        sep = QFrame()
        sep.setObjectName("separator")
        sep.setFrameShape(QFrame.Shape.HLine)
        lay.addWidget(sep)

        lbl_addr_caption = QLabel("Адрес:")
        lbl_addr_caption.setObjectName("panelCaption")
        lay.addWidget(lbl_addr_caption)

        self.lbl_address = QLabel("")
        self.lbl_address.setObjectName("panelReadOnly")
        self.lbl_address.setWordWrap(True)
        self.lbl_address.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        lay.addWidget(self.lbl_address)

        lbl_cur_caption = QLabel("Текущий номер:")
        lbl_cur_caption.setObjectName("panelCaption")
        lay.addWidget(lbl_cur_caption)

        self.lbl_current = QLabel("")
        self.lbl_current.setObjectName("panelHighlight")
        lay.addWidget(self.lbl_current)

        lbl_new_caption = QLabel("Новый номер:")
        lbl_new_caption.setObjectName("panelCaption")
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
        self.btn_save.setFixedHeight(30)
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
        # Обновляем отображение текущего номера и закрываем панель
        self.lbl_current.setText(new_val)
        self._on_close()

    def _on_close(self) -> None:
        self.clear()
        self.closed.emit()


# ─────────────────────────── Страница ─────────────────────────────────────

class PreviewGeneralPage(QWidget):
    """Предпросмотр и генерация файла «Общие маршруты»."""

    go_back         = pyqtSignal()
    go_home         = pyqtSignal()   # Переход на главную (dashboard)
    go_dept_preview = pyqtSignal()   # Переход к маршрутам по отделам
    go_settings     = pyqtSignal()   # Переход к настройкам Шт
    go_clear_routes = pyqtSignal()   # Очистить маршруты и вернуться на главную

    def __init__(self, app_state: dict):
        super().__init__()
        self.app_state       = app_state
        self._display_mode   = _DISPLAY_ADDR  # по умолчанию только № маршрута и адрес
        self._filter_product = ""
        self._search_text    = ""
        self._rendering      = False
        self._render_pending = False
        self._sort_asc       = True   # True = по возрастанию (по умолчанию)
        self._column_widths  = None   # пользовательские ширины столбцов в текущей сессии

        self._render_thread: QThread | None = None
        self._render_worker: RenderWorker | None = None

        # Debounce для поиска
        self._search_timer = QTimer(self)
        self._search_timer.setSingleShot(True)
        self._search_timer.timeout.connect(self._render_table)

        self._build_ui()

        # После построения UI разрешаем пользователю менять ширину столбцов мышью (как в Excel)
        try:
            hdr = self.table.horizontalHeader()
            for col in range(self._model.columnCount()):
                hdr.setSectionResizeMode(col, QHeaderView.ResizeMode.Interactive)
            hdr.sectionResized.connect(self._on_section_resized)
        except Exception:
            pass

    # ─────────────────────────── Построение UI ────────────────────────────

    def _build_ui(self) -> None:
        content = QWidget()
        content.setMinimumHeight(480)
        root_lay = QVBoxLayout(content)
        root_lay.setContentsMargins(20, 16, 20, 16)
        root_lay.setSpacing(12)

        # Заголовок
        h_row = QHBoxLayout()
        btn_back = QPushButton("← Назад")
        btn_back.setObjectName("btnBack")
        btn_back.clicked.connect(self.go_back.emit)
        h_row.addWidget(btn_back)
        btn_home = QPushButton("На главную")
        btn_home.setObjectName("btnSecondary")
        btn_home.clicked.connect(self.go_home.emit)
        h_row.addWidget(btn_home)
        self.btn_dept_preview = QPushButton("Маршруты по отделам")
        self.btn_dept_preview.setMinimumWidth(160)
        self.btn_dept_preview.setObjectName("btnSecondary")
        self.btn_dept_preview.clicked.connect(self.go_dept_preview.emit)
        self.btn_dept_preview.setEnabled(False)
        h_row.addWidget(self.btn_dept_preview)

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
        self.lbl_no_num.setWordWrap(True)
        h_row.addWidget(self.lbl_no_num)

        root_lay.addLayout(h_row)

        # Баннер непривязанных продуктов (4A)
        self.banner_unassigned = QFrame()
        self.banner_unassigned.setObjectName("bannerWarning")
        self.banner_unassigned.setVisible(False)
        banner_lay = QHBoxLayout(self.banner_unassigned)
        banner_lay.setContentsMargins(12, 8, 12, 8)
        self.lbl_banner = QLabel("")
        self.lbl_banner.setWordWrap(True)
        banner_lay.addWidget(self.lbl_banner)
        self.btn_banner_depts = QPushButton("Открыть Отделы и продукты")
        self.btn_banner_depts.setObjectName("btnPrimary")
        self.btn_banner_depts.clicked.connect(self._on_banner_open_departments)
        banner_lay.addWidget(self.btn_banner_depts)
        self.btn_banner_products = QPushButton("Справочник продуктов")
        self.btn_banner_products.setObjectName("btnSecondary")
        self.btn_banner_products.clicked.connect(self._on_banner_open_products)
        banner_lay.addWidget(self.btn_banner_products)
        root_lay.addWidget(self.banner_unassigned)

        # Панель фильтров (компактная)
        filter_card = QFrame()
        filter_card.setObjectName("card")
        filter_lay = QHBoxLayout(filter_card)
        filter_lay.setContentsMargins(12, 8, 12, 8)
        filter_lay.setSpacing(8)

        self.le_search = QLineEdit()
        self.le_search.setPlaceholderText("Поиск по адресу или номеру маршрута...")
        self.le_search.setClearButtonEnabled(True)
        self.le_search.setMinimumWidth(240)
        self.le_search.textChanged.connect(self._on_search_changed)
        filter_lay.addWidget(self.le_search)
        sc_search = QShortcut(QKeySequence("Ctrl+F"), self)
        sc_search.activated.connect(self.le_search.setFocus)

        filter_lay.addWidget(QLabel("Продукт:"))
        self.combo_product = QComboBox()
        self.combo_product.setMinimumWidth(180)
        self.combo_product.addItem("Все продукты", "")
        self.combo_product.currentIndexChanged.connect(self._on_product_filter_changed)
        filter_lay.addWidget(self.combo_product)

        filter_lay.addWidget(QLabel("Фильтр:"))
        self.combo_display = QComboBox()
        self.combo_display.setMinimumWidth(160)
        self.combo_display.addItem("Только № и адрес", _DISPLAY_ADDR)
        self.combo_display.addItem("Полностью",        _DISPLAY_FULL)
        self.combo_display.addItem("Только продукт",   _DISPLAY_PRODUCT)
        self.combo_display.setCurrentIndex(0)  # по умолчанию — только № и адрес
        self.combo_display.currentIndexChanged.connect(self._on_display_changed)
        filter_lay.addWidget(self.combo_display)

        # Кнопка сортировки
        self.btn_sort = QPushButton("↑ По возрастанию")
        self.btn_sort.setObjectName("btnSecondary")
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
        self.table.setObjectName("routesTable")
        self.table.setItemDelegate(RoutesTableDelegate(self.table))
        self.table.setModel(self._model)
        hdr = self.table.horizontalHeader()
        hdr.setMinimumSectionSize(90)
        # Широкие столбцы по умолчанию — меньше пустого места
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        hdr.resizeSection(0, 130)
        hdr.resizeSection(1, 420)
        hdr.resizeSection(2, 100)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.clicked.connect(self._on_cell_clicked)
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._on_context_menu)
        self._table_font_size = 11
        self.table.setFont(QFont("", self._table_font_size))
        self.table.installEventFilter(self)
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
        self.lbl_excluded.setObjectName("hintLabel")
        bottom_row.addWidget(self.lbl_excluded)
        bottom_row.addStretch()

        self.btn_generate = QPushButton("Создать файл «Общие маршруты»")
        self.btn_generate.setMinimumWidth(220)
        self.btn_generate.setObjectName("btnPrimary")
        self.btn_generate.setFixedHeight(32)
        self.btn_generate.clicked.connect(self._on_generate)
        bottom_row.addWidget(self.btn_generate)

        root_lay.addLayout(bottom_row)

        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.setVisible(False)
        root_lay.addWidget(self.progress)

        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setWidget(content)
        main_lay = QVBoxLayout(self)
        main_lay.setContentsMargins(0, 0, 0, 0)
        main_lay.addWidget(scroll)

        shortcut_escape = QShortcut(QKeySequence(Qt.Key.Key_Escape), self)
        shortcut_escape.activated.connect(self._on_escape_key)

    def _on_escape_key(self) -> None:
        """Закрыть панель редактирования номера маршрута по Escape."""
        if self.edit_panel.isVisible():
            self.edit_panel._on_close()

    def eventFilter(self, obj, event):
        """Ctrl + колёсико мыши — масштаб текста в таблице предпросмотра."""
        if obj == self.table and event.type() == QEvent.Type.Wheel:
            if QApplication.keyboardModifiers() & Qt.KeyboardModifier.ControlModifier:
                delta = event.angleDelta().y()
                step = 1 if delta > 0 else -1
                self._table_font_size = max(8, min(24, self._table_font_size + step))
                self._model.set_font_size(self._table_font_size)
                self.table.setFont(QFont("", self._table_font_size))
                self._model.emit_data_changed()
                return True
        return super().eventFilter(obj, event)

    # ─────────────────────────── Обновление ───────────────────────────────

    def _check_unassigned_products(self) -> list[str]:
        """Возвращает список продуктов без отдела из текущих маршрутов."""
        routes = self.app_state.get("filteredRoutes", [])
        products = data_store.get_ref("products") or []
        aliases = data_store.get_aliases()
        assigned = {p["name"] for p in products if p.get("deptKey")}
        unassigned: set[str] = set()
        for r in routes:
            if r.get("excluded"):
                continue
            for p in r.get("products", []):
                canonical = aliases.get(p["name"], p["name"])
                if canonical not in assigned:
                    unassigned.add(p["name"])
        return sorted(unassigned)

    def _on_banner_open_departments(self) -> None:
        from ui.pages import departments_page as dept_mod
        dept_mod.open_modal(self.window(), self.app_state)
        self.refresh()

    def _on_banner_open_products(self) -> None:
        from ui.pages.products_page import open_modal as open_products
        open_products(self.window(), self.app_state)
        self.refresh()

    def _get_routes_date_str(self) -> str:
        """Дата из app_state (задаётся при добавлении файлов) или завтра."""
        s = self.app_state.get("routesDate")
        if s:
            try:
                parts = s.split(".")
                if len(parts) == 3:
                    return s
            except (ValueError, TypeError):
                pass
        tomorrow = date.today() + timedelta(days=1)
        return f"{tomorrow.day:02d}.{tomorrow.month:02d}.{tomorrow.year}"

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
        has_routes = bool([r for r in self.app_state.get("filteredRoutes", []) if not r.get("excluded")])
        self.btn_dept_preview.setEnabled(has_routes)

        # Баннер и блокировка кнопки при непривязанных продуктах (4A)
        unassigned = self._check_unassigned_products()
        if unassigned:
            names_str = ", ".join(unassigned[:8])
            if len(unassigned) > 8:
                names_str += f" и ещё {len(unassigned) - 8}"
            self.lbl_banner.setText(
                f"⚠ {len(unassigned)} продукт(ов) без отдела: {names_str}. "
                "Сначала привяжите все продукты к отделам."
            )
            self.banner_unassigned.setVisible(True)
            self.btn_generate.setEnabled(False)
            self.btn_generate.setToolTip("Сначала привяжите все продукты к отделам")
        else:
            self.banner_unassigned.setVisible(False)
            self.btn_generate.setEnabled(True)
            self.btn_generate.setToolTip("")

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

        # Стартовые ширины столбцов — широкие по умолчанию; далее сохраняем пользовательские
        if self._column_widths is None:
            for i, w in enumerate([130, 420, 100, -1]):
                if i < self._model.columnCount() and w > 0:
                    self.table.setColumnWidth(i, w)
            self._column_widths = [
                self.table.columnWidth(i) for i in range(self._model.columnCount())
            ]
        else:
            # Восстанавливаем пользовательские ширины
            for i, w in enumerate(self._column_widths):
                if i < self._model.columnCount() and w > 0:
                    self.table.setColumnWidth(i, w)

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

        # Кнопка «Маршруты по отделам» активна, если есть данные о маршрутах
        has_routes = bool(all_routes) and sum(1 for r in all_routes if not r.get("excluded")) > 0
        self.btn_dept_preview.setEnabled(has_routes)

        log.debug("_on_render_done done")

        if self._render_pending:
            log.debug("render_pending=True, запускаем ещё раз")
            self._start_render_worker()

    def _on_section_resized(self, logical_index: int, _old: int, new: int) -> None:
        """Запоминаем ширины столбцов при ручном изменении пользователем."""
        if logical_index < 0 or new <= 0:
            return
        if self._column_widths is None:
            self._column_widths = [
                self.table.columnWidth(i) for i in range(self._model.columnCount())
            ]
            return
        if logical_index >= len(self._column_widths):
            self._column_widths.extend(
                self.table.columnWidth(i)
                for i in range(len(self._column_widths), self._model.columnCount())
            )
        self._column_widths[logical_index] = new

    # ─────────────────────────── Фильтры ──────────────────────────────────

    def _on_search_changed(self, text: str) -> None:
        self._search_text = text
        self._search_timer.start(200)

    def _on_product_filter_changed(self) -> None:
        self._filter_product = self.combo_product.currentData() or ""
        self._render_table()

    def _on_display_changed(self) -> None:
        self._display_mode = self.combo_display.currentData() or _DISPLAY_ADDR
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
        self.combo_display.setCurrentIndex(0)  # Только № и адрес
        self.combo_display.blockSignals(False)
        self._search_text    = ""
        self._filter_product = ""
        self._display_mode   = _DISPLAY_ADDR
        self._sort_asc       = True
        self.btn_sort.setText("↑ По возрастанию")
        self.app_state["sortAsc"] = True
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

        all_routes = self.app_state.get("filteredRoutes", [])
        route_address = route.get("address", "")
        # Ищем маршрут по адресу (стабильный ключ; route_ref может быть устаревшим)
        found_idx = next(
            (i for i, r in enumerate(all_routes) if r.get("address", "") == route_address),
            -1
        )
        if found_idx == -1:
            found_idx = next((i for i, r in enumerate(all_routes) if r is route), -1)
        if found_idx == -1:
            log.warning("route_num_saved: маршрут не найден в filteredRoutes (address=%r)",
                        route_address[:50] if route_address else "")
            return
        route = all_routes[found_idx]
        route["routeNum"] = new_val
        log.debug("route_num_saved: '%s' -> '%s', filteredRoutes[%d]", old_num, new_val, found_idx)

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
        """Обновляет метку: красный фон если есть неопределённые, зелёный если все определены."""
        self.lbl_no_num.setVisible(True)
        if no_num_count > 0:
            self.lbl_no_num.setText(f"Маршруты не определены: {no_num_count}")
            self.lbl_no_num.setStyleSheet(
                "background-color: #FEE2E2; color: #DC2626; border-radius: 12px; "
                "padding: 4px 10px; font-size: 10px; font-weight: 600;"
            )
        else:
            self.lbl_no_num.setText("Все маршруты определены")
            self.lbl_no_num.setStyleSheet(
                "background-color: #DCFCE7; color: #4BD08B; border-radius: 12px; "
                "padding: 4px 10px; font-size: 10px; font-weight: 600;"
            )

    def _on_panel_closed(self) -> None:
        self.table.clearSelection()

    # ─────────────────────────── Контекстное меню ─────────────────────────

    def _get_selected_route_refs(self) -> list:
        """Возвращает уникальные route_ref из выбранных строк."""
        indexes = self.table.selectionModel().selectedRows()
        seen: set = set()
        result: list = []
        for idx in indexes:
            rd = self._model.get_row(idx.row())
            if rd is None:
                continue
            ref = rd.get("route_ref")
            if ref is not None and id(ref) not in seen:
                seen.add(id(ref))
                result.append(ref)
        return result

    def _on_context_menu(self, pos) -> None:
        index = self.table.indexAt(pos)
        if not index.isValid():
            return
        rd = self._model.get_row(index.row())
        if rd is None:
            return

        selected_refs = self._get_selected_route_refs()
        if not selected_refs:
            return

        menu = QMenu(self)
        act_delete = None
        act_exclude = None
        act_toggle = None
        act_edit = None

        if len(selected_refs) > 1:
            act_delete = menu.addAction(f"Удалить выбранные ({len(selected_refs)} маршрутов)")
            excluded_count = sum(1 for r in selected_refs if r.get("excluded"))
            if excluded_count == 0:
                act_exclude = menu.addAction(f"Исключить выбранные ({len(selected_refs)})")
            else:
                act_exclude = menu.addAction(f"Включить выбранные ({len(selected_refs)})")
        else:
            route = selected_refs[0]
            if route.get("excluded"):
                act_toggle = menu.addAction("Включить маршрут")
            else:
                act_toggle = menu.addAction("Исключить маршрут")
            act_edit = menu.addAction("Изменить номер маршрута...")
            act_delete = menu.addAction("Удалить маршрут")

        action = menu.exec(self.table.viewport().mapToGlobal(pos))

        if action == act_delete:
            self._delete_routes(selected_refs)
        elif action == act_exclude:
            for r in selected_refs:
                r["excluded"] = not r.get("excluded", False)
            self._render_table()
        elif action == act_toggle:
            route["excluded"] = not route.get("excluded", False)
            self._render_table()
        elif action == act_edit:
            self.edit_panel.load(route)

    def _delete_routes(self, route_refs: list) -> None:
        """Удаляет маршруты из app_state (routes и filteredRoutes)."""
        if not route_refs:
            return
        n = len(route_refs)
        msg = f"Удалить {n} маршрут(ов)?" if n > 1 else "Удалить этот маршрут?"
        reply = QMessageBox.question(
            self, "Удаление маршрутов", msg,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        addresses = {r.get("address", "") for r in route_refs}
        routes = [r for r in self.app_state.get("routes", []) if r.get("address", "") not in addresses]
        filtered = [r for r in self.app_state.get("filteredRoutes", []) if r.get("address", "") not in addresses]
        self.app_state["routes"] = routes
        self.app_state["filteredRoutes"] = filtered
        self.edit_panel.clear()
        self._render_table()
        if hasattr(self.app_state.get("set_status"), "__call__"):
            self.app_state["set_status"](f"Удалено маршрутов: {n}")

    # ─────────────────────────── Генерация ────────────────────────────────────

    def _on_generate(self) -> None:
        routes        = self.app_state.get("filteredRoutes", [])
        active_routes = [r for r in routes if not r.get("excluded")]

        if not active_routes:
            total = len(routes)
            if total > 0:
                msg = "Все маршруты исключены. Снимите исключение (правый клик → Исключить/Включить) или добавьте маршруты."
            else:
                msg = "Нет маршрутов для создания файла."
            QMessageBox.warning(self, "Нет маршрутов", msg)
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

        file_type    = self.app_state.get("fileType", "main")
        base_dir = self.app_state.get("saveDir") or data_store.get_desktop_path()
        chosen_base = QFileDialog.getExistingDirectory(
            self,
            "Выберите папку для сохранения маршрутов",
            base_dir,
        )
        if not chosen_base:
            return
        self.app_state["saveDir"] = chosen_base
        date_str = self._get_routes_date_str()
        self.app_state["routesDate"] = date_str
        save_path = excel_generator.get_general_routes_path(chosen_base, file_type, date_str)
        os.makedirs(os.path.dirname(save_path), exist_ok=True)

        prod_map = data_store.get_products_map()
        self.btn_generate.setEnabled(False)
        self.progress.setVisible(True)

        self._gen_thread = QThread(self)
        self._gen_worker = GenerateWorker(
            active_routes, file_type, save_path, prod_map, self._sort_asc, date_str=date_str
        )
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
        if hasattr(self.app_state.get("set_status"), "__call__"):
            self.app_state["set_status"](f"Файл сохранён: {path}")
        file_type = self.app_state.get("fileType", "main")
        data_store.save_last_routes(
            file_type,
            self.app_state.get("routes", []),
            self.app_state.get("uniqueProducts", []),
            self.app_state.get("filteredRoutes", []),
            route_category=self.app_state.get("routeCategory"),
        )
        # Обновляем отчет по шт для папки дня.
        try:
            day_dir = os.path.dirname(os.path.dirname(path))
            date_str = self.app_state.get("routesDate") or excel_generator.get_routes_date_str()
            report_path = os.path.join(day_dir, f"Отчет Шт {date_str}.xls")
            main_blob = data_store.get_last_routes("main") or {}
            inc_blob = data_store.get_last_routes("increase") or {}
            excel_generator.generate_pcs_compare_report(
                report_path,
                main_blob.get("filteredRoutes") or main_blob.get("routes") or [],
                inc_blob.get("filteredRoutes") or inc_blob.get("routes") or [],
                data_store.get_ref("products") or [],
            )
        except Exception as exc:
            log.warning("Не удалось обновить отчет по Шт: %s", exc)
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
        date_str = self._get_routes_date_str()
        folder_name = f"Этикетки на {date_str}"
        out_dir = os.path.join(base_dir, folder_name)
        os.makedirs(out_dir, exist_ok=True)
        file_type = self.app_state.get("fileType", "main")
        departments_ref = data_store.get_ref("departments") or []
        self._labels_out_dir = out_dir
        self.btn_labels.setEnabled(False)
        self.progress.setVisible(True)
        self._labels_thread = QThread(self)
        self._labels_worker = TemplateLabelsWorker(
            routes, out_dir, file_type, products_ref, departments_ref
        )
        self._labels_worker.moveToThread(self._labels_thread)
        self._labels_thread.started.connect(self._labels_worker.run)
        self._labels_worker.finished.connect(self._on_labels_templates_done)
        self._labels_worker.error.connect(self._on_labels_templates_error)
        self._labels_worker.finished.connect(self._labels_thread.quit)
        self._labels_worker.error.connect(self._labels_thread.quit)
        self._labels_thread.finished.connect(self._labels_worker.deleteLater)
        self._labels_thread.finished.connect(self._labels_thread.deleteLater)
        self._labels_thread.start()

    def _on_labels_templates_done(self, created: list[str]) -> None:
        self.progress.setVisible(False)
        self.btn_labels.setEnabled(True)
        if created:
            QMessageBox.information(self, "Готово", f"Создано файлов: {len(created)}\n\n{self._labels_out_dir}")
            reply = QMessageBox.question(
                self,
                "Открыть безопасно",
                "Открыть первый созданный файл через безопасную локальную копию?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes,
            )
            if reply == QMessageBox.StandardButton.Yes:
                try:
                    open_excel_file_safely(created[0])
                except Exception as exc:
                    QMessageBox.warning(self, "Не удалось открыть файл", str(exc))
        else:
            QMessageBox.information(self, "Нет файлов", "Нет этикеток для создания.")

    def _on_labels_templates_error(self, msg: str) -> None:
        self.progress.setVisible(False)
        self.btn_labels.setEnabled(True)
        QMessageBox.critical(self, "Ошибка", msg)
