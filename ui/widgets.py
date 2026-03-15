"""
ui/widgets.py — Общие переиспользуемые виджеты.

CommitLineEdit — поле ввода с сохранением по Enter/потере фокуса.
HintIconButton — кнопка «!»: при наведении — краткая подсказка, по клику — подробная инструкция.
message_plain — QMessageBox с PlainText, чтобы переносы строк (\\n) отображались.
ToggleSwitch — iOS/Apple-style анимированный переключатель, замена QCheckBox.
ReplacementDiagramWidget — визуальная схема замены продукта (A → B или A → B + C).
"""
from __future__ import annotations

from PyQt6.QtWidgets import (
    QLineEdit, QPushButton, QMessageBox, QListWidget, QListWidgetItem,
    QFrame, QVBoxLayout, QHBoxLayout, QAbstractButton, QSizePolicy,
    QWidget, QLabel, QComboBox, QSlider,
)
from PyQt6.QtCore import pyqtSignal, pyqtProperty, QTimer, Qt, QEvent, QObject, QPropertyAnimation, QEasingCurve, QSize, QPoint
from PyQt6.QtGui import QPainter, QColor, QPen, QPolygon


def message_plain(parent, title: str, text: str, icon=QMessageBox.Icon.Information, buttons=QMessageBox.StandardButton.Ok):
    """Показывает QMessageBox с PlainText, чтобы переносы строк (\\n) в text отображались."""
    mb = QMessageBox(parent)
    mb.setWindowTitle(title)
    mb.setText(text)
    mb.setTextFormat(Qt.TextFormat.PlainText)
    mb.setIcon(icon)
    mb.setStandardButtons(buttons)
    return mb.exec()


class SearchableList(QFrame):
    """
    Поисковый список:
      - первая строка — пустое поле поиска;
      - ниже — QListWidget (одиночный или множественный выбор);
      - справа — кнопка-иконка для скрытия списка.
    """

    selection_changed = pyqtSignal()
    hide_requested = pyqtSignal()

    def __init__(self, parent=None, multi_select: bool = False):
        super().__init__(parent)
        self.setObjectName("searchableList")
        self.setFrameShape(QFrame.Shape.StyledPanel)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(4)

        top = QHBoxLayout()
        top.setContentsMargins(0, 0, 0, 0)
        top.setSpacing(4)

        self.edit_search = QLineEdit()
        self.edit_search.setPlaceholderText("Поиск...")
        top.addWidget(self.edit_search)

        self.btn_toggle = QPushButton("˄")
        self.btn_toggle.setObjectName("listToggleButton")
        self.btn_toggle.setFixedWidth(28)
        self.btn_toggle.setToolTip("Свернуть или развернуть список")
        self.btn_toggle.clicked.connect(self._on_hide_clicked)
        top.addWidget(self.btn_toggle)

        lay.addLayout(top)

        self.list = QListWidget()
        self.list.setAlternatingRowColors(True)
        self.list.itemSelectionChanged.connect(self.selection_changed)
        if multi_select:
            self.list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        else:
            self.list.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
        lay.addWidget(self.list)

        self.edit_search.textChanged.connect(self._apply_filter)

    def set_items(self, texts: list[str]) -> None:
        self.list.clear()
        for t in texts:
            self.list.addItem(QListWidgetItem(t))

    def selected_texts(self) -> list[str]:
        return [i.text() for i in self.list.selectedItems()]

    def clear_selection(self) -> None:
        self.list.clearSelection()

    def _apply_filter(self, text: str) -> None:
        pattern = (text or "").strip().lower()
        for row in range(self.list.count()):
            item = self.list.item(row)
            visible = not pattern or pattern in item.text().lower()
            item.setHidden(not visible)

    def _on_hide_clicked(self) -> None:
        self.hide_requested.emit()


def hint_icon_button(parent, short_text: str, long_text: str, title: str = "Инструкция"):
    """
    Создаёт кнопку «!» для подсказок: при наведении — краткое описание (tooltip),
    при нажатии — подробная инструкция по пунктам в отдельном окне.

    Args:
        parent: родительский виджет
        short_text: краткий текст при наведении
        long_text: подробная инструкция по пунктам (при клике)
        title: заголовок окна с инструкцией
    Returns:
        QPushButton с objectName "pageHintIcon"
    """
    btn = QPushButton("!", parent)
    btn.setObjectName("pageHintIcon")
    btn.setToolTip(short_text)
    btn.setCursor(Qt.CursorShape.PointingHandCursor)
    if long_text.strip():
        def _on_click():
            message_plain(parent, title, long_text.strip())
        btn.clicked.connect(_on_click)
    return btn


class ToggleSwitch(QAbstractButton):
    """iOS/Apple-style анимированный переключатель — полная замена QCheckBox.

    API совместим с QCheckBox:
      - сигнал stateChanged(int): 2 = включён, 0 = выключен
      - isChecked() / setChecked(bool)
      - toggled(bool) — унаследован от QAbstractButton
    """

    stateChanged = pyqtSignal(int)

    _TRACK_ON  = "#34C759"   # Apple green
    _TRACK_OFF = "#E5E5EA"   # Apple light grey
    _THUMB     = "#FFFFFF"
    _SHADOW    = "#00000030"
    _DURATION  = 180         # ms

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setCheckable(True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)

        self._pos: float = 0.0

        self._anim = QPropertyAnimation(self, b"_anim_pos", self)
        self._anim.setDuration(self._DURATION)
        self._anim.setEasingCurve(QEasingCurve.Type.InOutCubic)

        self.toggled.connect(self._on_toggled)

    # ── property for animation ──────────────────────────────────────────
    def _get_anim_pos(self) -> float:
        return self._pos

    def _set_anim_pos(self, value: float) -> None:
        self._pos = value
        self.update()

    _anim_pos = pyqtProperty(float, _get_anim_pos, _set_anim_pos)

    # ── size ─────────────────────────────────────────────────────────────
    def sizeHint(self) -> QSize:
        return QSize(48, 28)

    def minimumSizeHint(self) -> QSize:
        return self.sizeHint()

    # ── state ─────────────────────────────────────────────────────────────
    def setChecked(self, checked: bool) -> None:
        super().setChecked(checked)
        self._pos = 1.0 if checked else 0.0
        self.update()

    def _on_toggled(self, checked: bool) -> None:
        self._anim.stop()
        self._anim.setStartValue(self._pos)
        self._anim.setEndValue(1.0 if checked else 0.0)
        self._anim.start()
        self.stateChanged.emit(2 if checked else 0)

    # ── paint ─────────────────────────────────────────────────────────────
    def paintEvent(self, _event) -> None:
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)

        w, h = self.width(), self.height()
        t = self._pos  # 0.0 → off, 1.0 → on

        # Interpolate track colour
        c_off = QColor(self._TRACK_OFF)
        c_on  = QColor(self._TRACK_ON)
        r = int(c_off.red()   + t * (c_on.red()   - c_off.red()))
        g = int(c_off.green() + t * (c_on.green() - c_off.green()))
        b = int(c_off.blue()  + t * (c_on.blue()  - c_off.blue()))
        track_color = QColor(r, g, b)

        # Track
        p.setBrush(track_color)
        p.setPen(Qt.PenStyle.NoPen)
        p.drawRoundedRect(0, 0, w, h, h / 2, h / 2)

        # Thumb shadow ring
        margin = 2
        thumb_d = h - 2 * margin
        thumb_x = int(margin + t * (w - thumb_d - 2 * margin))
        shadow_pen = QPen(QColor(self._SHADOW))
        shadow_pen.setWidth(1)
        p.setPen(shadow_pen)
        p.setBrush(QColor(self._THUMB))
        p.drawEllipse(thumb_x, margin, thumb_d, thumb_d)

        p.end()


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


# ─────────────────────────── ReplacementDiagramWidget ───────────────────────

def _make_replacement_product_frame(min_width: int = 140) -> tuple[QFrame, QVBoxLayout]:
    """Создаёт стилизованный фрейм для продукта в схеме замены."""
    f = QFrame()
    f.setObjectName("replacementProductCard")
    f.setMinimumWidth(min_width)
    f.setMinimumHeight(56)
    f.setFrameStyle(QFrame.Shape.StyledPanel | QFrame.Shadow.Raised)
    lay = QVBoxLayout(f)
    lay.setContentsMargins(8, 6, 8, 6)
    lay.setSpacing(2)
    return f, lay


class _ArrowDiagramWidget(QWidget):
    """Виджет, рисующий стрелки: одна или две (ветвление)."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._single = True
        self.setMinimumWidth(80)
        self.setMinimumHeight(100)
        self.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Expanding)

    def set_single(self, single: bool) -> None:
        self._single = single
        self.update()

    def paintEvent(self, event):
        super().paintEvent(event)
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        pen = QPen(QColor(100, 100, 100), 2)
        p.setPen(pen)
        p.setBrush(QColor(100, 100, 100))

        w, h = self.width(), self.height()
        left_x, right_x = 8, w - 8

        if self._single:
            mid_y = h // 2
            p.drawLine(left_x, mid_y, right_x - 12, mid_y)
            pts = QPolygon([QPoint(right_x - 12, mid_y), QPoint(right_x - 22, mid_y - 6), QPoint(right_x - 22, mid_y + 6)])
            p.drawPolygon(pts)
        else:
            top_y, bot_y = h // 3, 2 * h // 3
            branch_x = left_x + (right_x - left_x) // 3
            p.drawLine(left_x, h // 2, branch_x, h // 2)
            p.drawLine(branch_x, h // 2, branch_x, top_y)
            p.drawLine(branch_x, top_y, right_x - 12, top_y)
            pts = QPolygon([QPoint(right_x - 12, top_y), QPoint(right_x - 22, top_y - 6), QPoint(right_x - 22, top_y + 6)])
            p.drawPolygon(pts)
            p.drawLine(branch_x, h // 2, branch_x, bot_y)
            p.drawLine(branch_x, bot_y, right_x - 12, bot_y)
            pts = QPolygon([QPoint(right_x - 12, bot_y), QPoint(right_x - 22, bot_y - 6), QPoint(right_x - 22, bot_y + 6)])
            p.drawPolygon(pts)
        p.end()


class ReplacementDiagramWidget(QWidget):
    """Визуальная схема замены: [A] → [B] или [A] → [B] + [C] с ветвлением."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._build_ui()

    def _build_ui(self) -> None:
        main_lay = QVBoxLayout(self)
        main_lay.setContentsMargins(0, 0, 0, 0)
        main_lay.setSpacing(8)

        row = QHBoxLayout()
        row.setSpacing(0)

        self.frame_from, from_lay = _make_replacement_product_frame(min_width=150)
        self.combo_from = QComboBox()
        self.combo_from.setMinimumWidth(130)
        from_lay.addWidget(self.combo_from)
        self.lbl_from_qty = QLabel("")
        self.lbl_from_qty.setObjectName("replacementHint")
        from_lay.addWidget(self.lbl_from_qty)
        row.addWidget(self.frame_from)

        self.arrow_widget = _ArrowDiagramWidget(self)
        row.addWidget(self.arrow_widget, 0, Qt.AlignmentFlag.AlignCenter)

        self.right_container = QWidget()
        self.right_lay = QVBoxLayout(self.right_container)
        self.right_lay.setContentsMargins(0, 0, 0, 0)
        self.right_lay.setSpacing(12)

        self.frame_to1, to1_lay = _make_replacement_product_frame(min_width=140)
        self.combo_to1 = QComboBox()
        self.combo_to1.setMinimumWidth(120)
        to1_lay.addWidget(self.combo_to1)
        self.lbl_to1_pct = QLabel("")
        self.lbl_to1_pct.setObjectName("replacementHint")
        to1_lay.addWidget(self.lbl_to1_pct)
        self.right_lay.addWidget(self.frame_to1)

        self.frame_to2, to2_lay = _make_replacement_product_frame(min_width=140)
        self.combo_to2 = QComboBox()
        self.combo_to2.setMinimumWidth(120)
        to2_lay.addWidget(self.combo_to2)
        self.lbl_to2_pct = QLabel("")
        self.lbl_to2_pct.setObjectName("replacementHint")
        to2_lay.addWidget(self.lbl_to2_pct)
        self.right_lay.addWidget(self.frame_to2)
        self.frame_to2.setVisible(False)

        row.addWidget(self.right_container)
        main_lay.addLayout(row)

        self.btn_add_second = QPushButton("+ Добавить второй продукт")
        self.btn_add_second.setObjectName("btnSecondary")
        self.btn_add_second.setCheckable(True)
        self.btn_add_second.setChecked(False)

        self.slider_row = QWidget()
        slider_lay = QHBoxLayout(self.slider_row)
        slider_lay.setContentsMargins(0, 4, 0, 0)
        slider_lay.addWidget(QLabel("Распределение:"))
        self.slider_ratio = QSlider(Qt.Orientation.Horizontal)
        self.slider_ratio.setRange(0, 100)
        self.slider_ratio.setValue(50)
        self.slider_ratio.setMinimumWidth(120)
        slider_lay.addWidget(self.slider_ratio)
        self.lbl_ratio = QLabel("50% / 50%")
        self.lbl_ratio.setMinimumWidth(70)
        slider_lay.addWidget(self.lbl_ratio)
        slider_lay.addStretch()
        self.slider_row.setVisible(False)

        self.slider_ratio.valueChanged.connect(lambda v: self.lbl_ratio.setText(f"{v}% / {100 - v}%"))

        def _on_add_toggled(checked: bool) -> None:
            self.frame_to2.setVisible(checked)
            self.slider_row.setVisible(checked)
            self.arrow_widget.set_single(not checked)
            self.btn_add_second.setText("− Убрать второй продукт" if checked else "+ Добавить второй продукт")

        self.btn_add_second.toggled.connect(_on_add_toggled)
        main_lay.addWidget(self.btn_add_second)
        main_lay.addWidget(self.slider_row)
