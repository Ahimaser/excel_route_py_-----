"""
ui/widgets.py — Общие переиспользуемые виджеты.

CommitLineEdit — поле ввода с сохранением по Enter/потере фокуса.
HintIconButton — кнопка «!»: при наведении — краткая подсказка, по клику — подробная инструкция.
make_combo_searchable — при нажатии список отображается целиком; при вводе символов остаются только пункты с совпадением.
message_plain — QMessageBox с PlainText, чтобы переносы строк (\\n) отображались.
ToggleSwitch — iOS/Apple-style анимированный переключатель, замена QCheckBox.
"""
from __future__ import annotations

from PyQt6.QtWidgets import QLineEdit, QPushButton, QMessageBox, QComboBox, QListWidget, QListWidgetItem, QFrame, QVBoxLayout, QHBoxLayout, QAbstractButton, QSizePolicy
from PyQt6.QtCore import pyqtSignal, pyqtProperty, QTimer, Qt, QSortFilterProxyModel, QModelIndex, QEvent, QObject, QPropertyAnimation, QEasingCurve, QSize
from PyQt6.QtGui import QPainter, QColor, QPen


def message_plain(parent, title: str, text: str, icon=QMessageBox.Icon.Information, buttons=QMessageBox.StandardButton.Ok):
    """Показывает QMessageBox с PlainText, чтобы переносы строк (\\n) в text отображались."""
    mb = QMessageBox(parent)
    mb.setWindowTitle(title)
    mb.setText(text)
    mb.setTextFormat(Qt.TextFormat.PlainText)
    mb.setIcon(icon)
    mb.setStandardButtons(buttons)
    return mb.exec()


class _ComboFilterProxy(QSortFilterProxyModel):
    """Фильтр по вхождению подстроки в текст пункта (без учёта регистра)."""

    def filterAcceptsRow(self, source_row: int, source_parent: QModelIndex) -> bool:
        rx = self.filterRegularExpression()
        if not rx.pattern():
            return True
        src = self.sourceModel()
        if not src:
            return True
        idx = src.index(source_row, 0, source_parent)
        text = (src.data(idx, Qt.ItemDataRole.DisplayRole) or "")
        if not isinstance(text, str):
            text = str(text)
        return rx.match(text).hasMatch()


def make_combo_searchable(combo: QComboBox) -> None:
    """
    Список при нажатии отображается целиком; при вводе символов в поле
    в списке остаются только пункты, в которых есть введённая подстрока.
    """
    combo.setEditable(True)
    combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
    source = combo.model()
    old_index = combo.currentIndex()
    row_count = source.rowCount()
    proxy = _ComboFilterProxy(combo)
    proxy.setSourceModel(source)
    proxy.setFilterCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
    combo.setModel(proxy)
    if 0 <= old_index < row_count:
        combo.setCurrentIndex(old_index)

    def _apply_filter(text: str) -> None:
        t = (text or "").strip()
        if not t:
            proxy.setFilterRegularExpression("")
        else:
            import re
            from PyQt6.QtCore import QRegularExpression
            pattern = ".*" + re.escape(t) + ".*"
            proxy.setFilterRegularExpression(
                QRegularExpression(pattern, QRegularExpression.PatternOption.CaseInsensitiveOption)
            )

    le = combo.lineEdit()
    if le is not None:
        le.setPlaceholderText("Поиск...")

    class _FocusFilter(QObject):
        """При фокусе в поле сбрасываем фильтр, чтобы при открытии списка он отображался целиком."""

        def __init__(self, proxy_model: QSortFilterProxyModel, line_edit: QLineEdit, parent=None):
            super().__init__(parent)
            self._proxy = proxy_model
            self._line_edit = line_edit

        def eventFilter(self, obj, ev):
            if ev.type() == QEvent.Type.FocusIn:
                # При первом фокусе очищаем текст, чтобы поле было пустым для поиска
                self._line_edit.clear()
                self._proxy.setFilterRegularExpression("")
            return False

    if le is not None:
        le.installEventFilter(_FocusFilter(proxy, le, combo))
        le.textChanged.connect(_apply_filter)

    # Минимальный размер списка, чтобы пункты были видны и кликабельны (в т.ч. внутри QScrollArea)
    view = combo.view()
    view.setMinimumHeight(min(300, 80 + max(0, row_count) * 24))
    view.setMinimumWidth(max(220, combo.minimumSizeHint().width()))

    # После клика по комбобоксу через 50 ms popup уже открыт — поднимаем его поверх (обрезка в QScrollArea)
    def _raise_popup():
        popup = view.window()
        if popup and popup is not combo and popup.isVisible():
            flags = popup.windowFlags()
            if not (flags & Qt.WindowType.WindowStaysOnTopHint):
                popup.setWindowFlags(flags | Qt.WindowType.WindowStaysOnTopHint)

    class _ComboPopupRaiser(QObject):
        def eventFilter(self, obj, ev):
            # После клика даём Qt открыть popup и только затем поднимаем его поверх
            if obj is combo and ev.type() == QEvent.Type.MouseButtonPress:
                QTimer.singleShot(50, _raise_popup)
            return False
    combo.installEventFilter(_ComboPopupRaiser(combo))


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
        self.btn_toggle.setToolTip("Скрыть список")
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
