"""
ui/widgets.py — Общие переиспользуемые виджеты.

CommitLineEdit — поле ввода с сохранением по Enter/потере фокуса.
HintIconButton — кнопка «!»: при наведении — краткая подсказка, по клику — подробная инструкция.
make_combo_searchable — при нажатии список отображается целиком; при вводе символов остаются только пункты с совпадением.
"""
from __future__ import annotations

from PyQt6.QtWidgets import QLineEdit, QPushButton, QMessageBox, QComboBox
from PyQt6.QtCore import pyqtSignal, QTimer, Qt, QSortFilterProxyModel, QModelIndex, QEvent, QObject


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
    proxy = _ComboFilterProxy(combo)
    proxy.setSourceModel(source)
    proxy.setFilterCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
    combo.setModel(proxy)
    if 0 <= old_index < source.rowCount():
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

    class _FocusFilter(QObject):
        """При фокусе в поле сбрасываем фильтр, чтобы при открытии списка он отображался целиком."""

        def __init__(self, proxy_model: QSortFilterProxyModel, parent=None):
            super().__init__(parent)
            self._proxy = proxy_model

        def eventFilter(self, obj, ev):
            if ev.type() == QEvent.Type.FocusIn:
                self._proxy.setFilterRegularExpression("")
            return False

    le.installEventFilter(_FocusFilter(proxy, combo))
    le.textChanged.connect(_apply_filter)


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
            QMessageBox.information(parent, title, long_text.strip())
        btn.clicked.connect(_on_click)
    return btn


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
