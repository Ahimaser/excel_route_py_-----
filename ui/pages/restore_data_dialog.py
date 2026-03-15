"""
restore_data_dialog.py — Диалог восстановления данных из резервной копии store.json.
"""
from __future__ import annotations

from datetime import datetime

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QListWidget, QListWidgetItem, QMessageBox,
)
from PyQt6.QtCore import Qt

from core import data_store


def run_restore_data_dialog(parent) -> bool:
    """
    Открывает диалог выбора резервной копии и восстанавливает данные.
    Возвращает True если восстановление выполнено (рекомендуется перезапуск).
    """
    backups = data_store.list_backups()
    if not backups:
        QMessageBox.information(
            parent, "Резервные копии",
            "Резервные копии не найдены.\n\nОни создаются автоматически при каждом сохранении настроек."
        )
        return False

    dlg = QDialog(parent)
    dlg.setWindowTitle("Восстановить данные из резервной копии")
    lay = QVBoxLayout(dlg)

    lay.addWidget(QLabel("Выберите резервную копию для восстановления:"))
    list_widget = QListWidget()
    for idx, name, mtime in backups:
        dt = datetime.fromtimestamp(mtime).strftime("%d.%m.%Y %H:%M")
        item = QListWidgetItem(f"{name} — {dt}")
        item.setData(Qt.ItemDataRole.UserRole, idx)
        list_widget.addItem(item)
    list_widget.setCurrentRow(0)
    lay.addWidget(list_widget)

    btn_lay = QHBoxLayout()
    btn_lay.addStretch()
    btn_restore = QPushButton("Восстановить")
    btn_cancel = QPushButton("Отмена")
    btn_restore.clicked.connect(dlg.accept)
    btn_cancel.clicked.connect(dlg.reject)
    btn_lay.addWidget(btn_restore)
    btn_lay.addWidget(btn_cancel)
    lay.addLayout(btn_lay)

    if dlg.exec() != QDialog.DialogCode.Accepted:
        return False

    item = list_widget.currentItem()
    if not item:
        return False
    backup_index = item.data(Qt.ItemDataRole.UserRole)
    if backup_index is None:
        return False

    reply = QMessageBox.question(
        parent, "Подтверждение",
        "Восстановить данные из выбранной резервной копии?\n\n"
        "Рекомендуется перезапустить приложение после восстановления.",
        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        QMessageBox.StandardButton.No,
    )
    if reply != QMessageBox.StandardButton.Yes:
        return False

    if data_store.restore_from_backup(backup_index):
        QMessageBox.information(
            parent, "Готово",
            "Данные восстановлены.\n\nПерезапустите приложение для применения изменений."
        )
        return True
    QMessageBox.critical(
        parent, "Ошибка",
        "Не удалось восстановить данные из резервной копии."
    )
    return False
