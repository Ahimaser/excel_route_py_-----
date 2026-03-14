from __future__ import annotations

import os
import shutil
import tempfile
from pathlib import Path


def _strip_zone_identifier(path: str) -> None:
    if os.name != "nt":
        return
    try:
        os.remove(f"{os.path.abspath(path)}:Zone.Identifier")
    except Exception:
        pass


def open_excel_file_safely(path: str) -> str:
    """
    Открывает Excel-файл через локальную временную копию.
    Возвращает путь к временной копии.
    """
    src = os.path.abspath(path)
    if not os.path.isfile(src):
        raise FileNotFoundError(src)

    _strip_zone_identifier(src)
    temp_dir = tempfile.mkdtemp(prefix="excel_safe_open_")
    dst = os.path.join(temp_dir, Path(src).name)
    shutil.copyfile(src, dst)
    _strip_zone_identifier(dst)

    if os.name == "nt" and hasattr(os, "startfile"):
        os.startfile(dst)  # type: ignore[attr-defined]
    else:
        raise RuntimeError("Безопасное открытие поддерживается только на Windows.")
    return dst
