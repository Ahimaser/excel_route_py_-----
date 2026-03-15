"""Тесты для data_store: list_backups, restore_from_backup."""
import pytest
from pathlib import Path
from core import data_store


def test_list_backups_returns_list():
    """list_backups возвращает список (кортежей или пустой)."""
    result = data_store.list_backups()
    assert isinstance(result, list)
    for item in result:
        assert len(item) == 3
        idx, name, mtime = item
        assert isinstance(idx, int)
        assert isinstance(name, str)
        assert isinstance(mtime, (int, float))


def test_restore_from_backup_invalid_index():
    """restore_from_backup с несуществующим индексом возвращает False."""
    assert data_store.restore_from_backup(99) is False
