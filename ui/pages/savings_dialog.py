"""
savings_dialog.py — Окно «Экономия»: уменьшение количества по учреждениям/адресам.

Логика: список учреждений (группировка как в округлении), для каждого можно задать
% уменьшения продукта (только для ед. изм. не «шт»). Можно для всего учреждения
или для конкретного адреса. Приоритет: сначала экономия, потом округление.
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
    QDoubleSpinBox,
    QLineEdit,
    QScrollArea,
    QFrame,
    QWidget,
)

from core import data_store
from ui.widgets import ToggleSwitch, hint_icon_button


class SavingsDialog(QDialog):
    """Диалог настройки экономии по учреждениям и адресам."""

    def __init__(self, parent: QWidget | None, app_state: dict):
        super().__init__(parent)
        self._app_state = app_state
        routes = app_state.get("routes") or app_state.get("filteredRoutes") or []
        self._inst_addresses = data_store.get_institution_addresses_map(routes)
        self._settings = data_store.get_savings_settings()
        self._current_code: str | None = None

        self.setWindowTitle("Экономия")
        self.setMinimumSize(700, 500)
        self.resize(850, 550)
        self._build_ui()
        self._populate_institutions()
        self._update_summary()

    def _build_ui(self) -> None:
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(16)

        hint_row = QHBoxLayout()
        hint = QLabel(
            "Уменьшение количества продукта в % для учреждений и адресов. "
            "Применяется только к продуктам с ед. изм. не «шт». "
            "Приоритет: сначала экономия, затем округление."
        )
        hint.setObjectName("stepLabel")
        hint.setWordWrap(True)
        hint_row.addWidget(hint)
        hint_row.addWidget(hint_icon_button(
            self,
            "Экономия уменьшает количество до округления.",
            "Инструкция — Экономия\n\n"
            "1. Выберите учреждение слева — справа отобразятся его адреса.\n"
            "2. «Уменьшение для всего учреждения» — % применяется ко всем адресам.\n"
            "3. Для адреса можно задать свой % или исключить из учреждения.\n"
            "4. Приоритет: сначала экономия, затем округление Шт.",
            "Инструкция",
        ))
        hint_row.addStretch()
        lay.addLayout(hint_row)

        self.lbl_no_routes = QLabel("Нет учреждений в текущих маршрутах.")
        self.lbl_no_routes.setObjectName("hintLabel")
        self.lbl_no_routes.setVisible(False)
        lay.addWidget(self.lbl_no_routes)

        body = QHBoxLayout()
        body.setSpacing(16)

        left_card = QFrame()
        left_card.setObjectName("card")
        left_lay = QVBoxLayout(left_card)
        left_lay.setContentsMargins(16, 16, 16, 16)
        left_lay.setSpacing(10)
        left_lay.addWidget(QLabel("Учреждения"))
        self.search_inst = QLineEdit()
        self.search_inst.setPlaceholderText("Поиск учреждения или адреса")
        self.search_inst.textChanged.connect(self._apply_inst_filter)
        left_lay.addWidget(self.search_inst)
        self.lbl_inst_summary = QLabel()
        self.lbl_inst_summary.setObjectName("hintLabel")
        self.lbl_inst_summary.setWordWrap(True)
        left_lay.addWidget(self.lbl_inst_summary)
        self.inst_list = QListWidget()
        self.inst_list.setAlternatingRowColors(True)
        self.inst_list.currentItemChanged.connect(self._on_institution_selected)
        left_lay.addWidget(self.inst_list, 1)
        body.addWidget(left_card, 1)

        right_card = QFrame()
        right_card.setObjectName("card")
        right_lay = QVBoxLayout(right_card)
        right_lay.setContentsMargins(16, 16, 16, 16)
        right_lay.setSpacing(12)

        self.lbl_current_inst = QLabel("Выберите учреждение слева")
        self.lbl_current_inst.setObjectName("sectionTitle")
        self.lbl_current_inst.setWordWrap(True)
        right_lay.addWidget(self.lbl_current_inst)

        row_percent = QHBoxLayout()
        row_percent.addWidget(QLabel("Уменьшение для всего учреждения (%):"))
        self.spin_inst_percent = QDoubleSpinBox()
        self.spin_inst_percent.setRange(0, 100)
        self.spin_inst_percent.setDecimals(1)
        self.spin_inst_percent.setSuffix(" %")
        self.spin_inst_percent.setValue(0)
        self.spin_inst_percent.valueChanged.connect(self._on_inst_percent_changed)
        row_percent.addWidget(self.spin_inst_percent)
        row_percent.addStretch()
        right_lay.addLayout(row_percent)

        right_lay.addWidget(QLabel("Адреса (исключить из учреждения или задать свой %):"))
        self.search_addr = QLineEdit()
        self.search_addr.setPlaceholderText("Поиск адреса")
        self.search_addr.textChanged.connect(self._refresh_addr_panel)
        right_lay.addWidget(self.search_addr)

        self.addr_scroll = QScrollArea()
        self.addr_scroll.setWidgetResizable(True)
        self.addr_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.addr_container = QWidget()
        self.addr_lay = QVBoxLayout(self.addr_container)
        self.addr_lay.setContentsMargins(0, 0, 0, 0)
        self.addr_lay.setSpacing(8)
        self.addr_scroll.setWidget(self.addr_container)
        right_lay.addWidget(self.addr_scroll, 1)

        body.addWidget(right_card, 2)
        lay.addLayout(body, 1)

        btn_lay = QHBoxLayout()
        btn_lay.addStretch()
        btn_cancel = QPushButton("Отмена")
        btn_cancel.setObjectName("btnSecondary")
        btn_cancel.clicked.connect(self.reject)
        btn_lay.addWidget(btn_cancel)
        btn_ok = QPushButton("Сохранить")
        btn_ok.setObjectName("btnPrimary")
        btn_ok.clicked.connect(self._save)
        btn_lay.addWidget(btn_ok)
        lay.addLayout(btn_lay)

    def _populate_institutions(self) -> None:
        self.inst_list.clear()
        codes = sorted(self._inst_addresses.keys())
        if not codes:
            self.lbl_no_routes.setVisible(True)
            return
        self.lbl_no_routes.setVisible(False)
        for code in codes:
            addrs = self._inst_addresses.get(code, [])
            inst_cfg = self._settings.get("institutionSavings", {}).get(code, {})
            pct = inst_cfg.get("percent", 0) or 0
            item = QListWidgetItem(f"{code} — {len(addrs)} адр. ({pct}%)")
            item.setData(Qt.ItemDataRole.UserRole, code)
            self.inst_list.addItem(item)
        self._apply_inst_filter()

    def _apply_inst_filter(self) -> None:
        q = (self.search_inst.text() or "").strip().lower()
        for i in range(self.inst_list.count()):
            item = self.inst_list.item(i)
            code = item.data(Qt.ItemDataRole.UserRole) or ""
            addrs = self._inst_addresses.get(code, [])
            show = not q or q in code.lower() or any(q in a.lower() for a in addrs)
            item.setHidden(not show)

    def _on_institution_selected(self, cur: QListWidgetItem | None, _prev: QListWidgetItem | None = None) -> None:
        if not cur:
            self._current_code = None
            self.lbl_current_inst.setText("Выберите учреждение слева")
            self.spin_inst_percent.setValue(0)
            self._refresh_addr_panel()
            return
        code = cur.data(Qt.ItemDataRole.UserRole) or ""
        self._current_code = code
        addrs = self._inst_addresses.get(code, [])
        self.lbl_current_inst.setText(f"{code} ({len(addrs)} адресов)")
        inst_cfg = self._settings.get("institutionSavings", {}).get(code, {})
        self.spin_inst_percent.blockSignals(True)
        self.spin_inst_percent.setValue(float(inst_cfg.get("percent", 0) or 0))
        self.spin_inst_percent.blockSignals(False)
        self._refresh_addr_panel()

    def _refresh_addr_panel(self) -> None:
        for i in reversed(range(self.addr_lay.count())):
            w = self.addr_lay.takeAt(i).widget()
            if w:
                w.deleteLater()
        if not self._current_code:
            return
        addrs = self._inst_addresses.get(self._current_code, [])
        inst_cfg = self._settings.get("institutionSavings", {}).get(self._current_code, {})
        exclude = set(inst_cfg.get("excludeAddresses") or [])
        addr_savings = self._settings.get("addressSavings") or {}
        q = (self.search_addr.text() or "").strip().lower()
        for addr in addrs:
            if q and q not in addr.lower():
                continue
            row = QFrame()
            row.setObjectName("card")
            row_lay = QHBoxLayout(row)
            row_lay.setContentsMargins(12, 10, 12, 10)
            row_lay.setSpacing(12)

            addr_lbl = QLabel(addr[:60] + "…" if len(addr) > 60 else addr)
            addr_lbl.setWordWrap(True)
            addr_lbl.setObjectName("stepLabel")
            row_lay.addWidget(addr_lbl, 1)

            spin = QDoubleSpinBox()
            spin.setRange(0, 100)
            spin.setDecimals(1)
            spin.setSuffix(" %")
            addr_cfg = addr_savings.get(addr, {})
            spin.setValue(float(addr_cfg.get("percent", 0) or 0))
            spin.valueChanged.connect(self._make_percent_handler(addr))
            spin.setFixedWidth(80)
            row_lay.addWidget(spin)

            tog = ToggleSwitch()
            tog.setChecked(addr not in exclude)
            tog.setToolTip("Выкл. — исключить адрес из экономии учреждения")
            tog.stateChanged.connect(self._make_exclude_handler(addr))
            row_lay.addWidget(tog)

            self.addr_lay.addWidget(row)

    def _make_exclude_handler(self, addr: str):
        def handler(state: int) -> None:
            excluded = state != Qt.CheckState.Checked.value
            self._set_addr_excluded(addr, excluded)
        return handler

    def _make_percent_handler(self, addr: str):
        def handler(v: float) -> None:
            self._set_addr_percent(addr, v)
        return handler

    def _set_addr_excluded(self, addr: str, excluded: bool) -> None:
        inst_savings = dict(self._settings.get("institutionSavings") or {})
        if not self._current_code:
            return
        cfg = dict(inst_savings.get(self._current_code, {}))
        exc = list(cfg.get("excludeAddresses") or [])
        if excluded and addr not in exc:
            exc.append(addr)
        elif not excluded and addr in exc:
            exc.remove(addr)
        cfg["excludeAddresses"] = exc
        inst_savings[self._current_code] = cfg
        self._settings["institutionSavings"] = inst_savings
        self._update_summary()

    def _set_addr_percent(self, addr: str, percent: float) -> None:
        addr_savings = dict(self._settings.get("addressSavings") or {})
        if percent > 0:
            addr_savings[addr] = {"percent": percent}
        else:
            addr_savings.pop(addr, None)
        self._settings["addressSavings"] = addr_savings
        self._update_summary()

    def _on_inst_percent_changed(self, value: float) -> None:
        if not self._current_code:
            return
        inst_savings = dict(self._settings.get("institutionSavings") or {})
        cfg = dict(inst_savings.get(self._current_code, {}))
        cfg["percent"] = value
        inst_savings[self._current_code] = cfg
        self._settings["institutionSavings"] = inst_savings
        cur = self.inst_list.currentItem()
        if cur and cur.data(Qt.ItemDataRole.UserRole) == self._current_code:
            addrs = self._inst_addresses.get(self._current_code, [])
            cur.setText(f"{self._current_code} — {len(addrs)} адр. ({value}%)")
        self._update_summary()

    def _update_summary(self) -> None:
        inst_count = sum(1 for c, cfg in (self._settings.get("institutionSavings") or {}).items()
                        if (cfg.get("percent") or 0) > 0)
        addr_count = len(self._settings.get("addressSavings") or {})
        self.lbl_inst_summary.setText(
            f"Учреждений с экономией: {inst_count}. Адресов с отдельным %: {addr_count}."
        )

    def _save(self) -> None:
        data_store.set_savings_settings(self._settings)
        self.accept()


def open_savings_dialog(parent: QWidget, app_state: dict) -> None:
    """Открывает диалог «Экономия»."""
    dlg = SavingsDialog(parent, app_state)
    dlg.exec()
