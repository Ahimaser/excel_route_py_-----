"""
data_store.py — JSON хранилище настроек приложения.

Оптимизации:
- Прямой доступ к данным без лишних deep-copy (get_ref/set_key)
- Атомарная запись через временный файл (нет риска повреждения)
- Ленивая загрузка (load при первом обращении)
- Отложенная запись через _dirty-флаг (flush только при реальных изменениях)
- Кэш пути к рабочему столу
"""
from __future__ import annotations

import json
import os
import copy
import tempfile
import uuid
from pathlib import Path
from typing import Any

APP_NAME = "ExcelRouteManager"

# ─────────────────────────── Дефолтные значения ───────────────────────────

DEFAULTS: dict[str, Any] = {
    "departments": [],
    "products": [],
    # product_aliases: {"вариант написания": "каноническое название"}
    # Используется парсером для автозамены при чтении файлов
    "product_aliases": {},
    # Новый формат столбцов шаблона:
    # { "field": "routeNumber"|"address"|"unit"|"quantity"|"pcs"|"productQty",
    #   "label": "Заголовок" (None = дефолт),
    #   "merged": False }
    # Для объединённого столбца (productQty):
    # { "field": "productQty", "label": None, "merged": True,
    #   "productName": "Название продукта" }
    "templates": [
        {
            "id": "template1",
            "name": "Шаблон 1 — Полный",
            "columns": [
                {"field": "routeNumber", "label": None, "merged": False},
                {"field": "address",     "label": None, "merged": False},
                {"field": "product",     "label": None, "merged": False},
                {"field": "unit",        "label": None, "merged": False},
                {"field": "quantity",    "label": None, "merged": False},
                {"field": "pcs",         "label": None, "merged": False},
            ],
            "deptKey": None,
        },
        {
            "id": "template2",
            "name": "Шаблон 2 — Компактный",
            "columns": [
                {"field": "routeNumber", "label": None, "merged": False},
                {"field": "address",     "label": None, "merged": False},
                {"field": "product",     "label": None, "merged": False},
                {"field": "quantity",    "label": None, "merged": False},
            ],
            "deptKey": None,
        },
    ],
    "settings": {
        "defaultSaveDir": None,
        "showPcsInPreview": True,
    },
}

# ─────────────────────────── Состояние модуля ─────────────────────────────

_data: dict[str, Any] | None = None
_path: Path | None = None
_dirty: bool = False
_desktop_cache: str | None = None


# ─────────────────────────── Внутренние утилиты ───────────────────────────

def _get_data_path() -> Path:
    if os.name == "nt":
        base = Path(os.environ.get("APPDATA", Path.home()))
    else:
        base = Path.home() / ".config"
    folder = base / APP_NAME
    folder.mkdir(parents=True, exist_ok=True)
    return folder / "store.json"


def _ensure_loaded() -> None:
    """Загружает данные из файла при первом обращении."""
    global _data, _path
    if _data is not None:
        return
    _path = _get_data_path()
    if _path.exists():
        try:
            with open(_path, "r", encoding="utf-8") as f:
                _data = json.load(f)
        except Exception:
            _data = {}
    else:
        _data = {}
    # Заполняем дефолтами для отсутствующих ключей
    for key, val in DEFAULTS.items():
        if key not in _data:
            _data[key] = copy.deepcopy(val)
    if not _data.get("templates"):
        _data["templates"] = copy.deepcopy(DEFAULTS["templates"])
    else:
        # Миграция: преобразуем старый формат столбцов (list[str]) в новый (list[dict])
        for tmpl in _data["templates"]:
            cols = tmpl.get("columns", [])
            if cols and isinstance(cols[0], str):
                tmpl["columns"] = [
                    {"field": c, "label": None, "merged": False} for c in cols
                ]
    if "product_aliases" not in _data:
        _data["product_aliases"] = {}


def _flush() -> None:
    """Атомарная запись данных на диск через временный файл."""
    global _dirty
    if not _dirty or _data is None or _path is None:
        return
    try:
        dir_ = _path.parent
        with tempfile.NamedTemporaryFile(
            "w", encoding="utf-8", dir=dir_, delete=False, suffix=".tmp"
        ) as tf:
            json.dump(_data, tf, ensure_ascii=False, indent=2)
            tmp_path = tf.name
        # Атомарная замена (работает на Windows и Linux)
        os.replace(tmp_path, _path)
        _dirty = False
    except Exception as e:
        print(f"[DataStore] flush error: {e}")


# ─────────────────────────── Публичный API ────────────────────────────────

def get(key: str) -> Any:
    """
    Возвращает глубокую копию значения по ключу.
    Используйте get_ref() для чтения без копирования (не изменяйте результат!).
    """
    _ensure_loaded()
    val = _data.get(key)
    if val is None:
        return None
    # Для простых типов копирование не нужно
    if isinstance(val, (str, int, float, bool)):
        return val
    return copy.deepcopy(val)


def get_ref(key: str) -> Any:
    """
    Возвращает прямую ссылку на данные (без копирования).
    ВНИМАНИЕ: не изменяйте возвращённый объект напрямую — используйте set_key().
    Используется для read-only операций в hot-path (рендер таблиц и т.п.).
    """
    _ensure_loaded()
    return _data.get(key)


def get_all() -> dict[str, Any]:
    """Возвращает глубокую копию всего хранилища."""
    _ensure_loaded()
    return copy.deepcopy(_data)


def set_key(key: str, value: Any) -> None:
    """Устанавливает значение и помечает хранилище как изменённое."""
    global _dirty
    _ensure_loaded()
    _data[key] = value
    _dirty = True
    _flush()


def update_product(name: str, **kwargs) -> bool:
    """
    Обновляет поля одного продукта по имени без перезаписи всего списка.
    Возвращает True если продукт найден и обновлён.
    """
    global _dirty
    _ensure_loaded()
    products: list = _data.get("products", [])
    for p in products:
        if p.get("name") == name:
            p.update(kwargs)
            _dirty = True
            _flush()
            return True
    return False


def get_desktop_path() -> str:
    """Возвращает путь к рабочему столу (кэшируется)."""
    global _desktop_cache
    if _desktop_cache is not None:
        return _desktop_cache
    desktop = Path.home() / "Desktop"
    _desktop_cache = str(desktop) if desktop.exists() else str(Path.home())
    return _desktop_cache


def get_products_map() -> dict[str, dict]:
    """
    Возвращает словарь {name: product_dict} без копирования.
    Используется в hot-path генерации и рендера таблиц.
    """
    _ensure_loaded()
    products = _data.get("products", [])
    return {p["name"]: p for p in products}


def get_aliases() -> dict[str, str]:
    """
    Возвращает прямую ссылку на словарь алиасов {вариант: каноническое}.
    Используется парсером для автозамены названий продуктов.
    """
    _ensure_loaded()
    return _data.get("product_aliases", {})


def set_alias(variant: str, canonical: str) -> None:
    """
    Добавляет или обновляет алиас: variant -> canonical.
    Если variant == canonical — удаляет алиас (отменяет связку).
    """
    global _dirty
    _ensure_loaded()
    aliases: dict = _data.setdefault("product_aliases", {})
    if variant == canonical:
        aliases.pop(variant, None)
    else:
        aliases[variant] = canonical
    _dirty = True
    _flush()


def remove_alias(variant: str) -> None:
    """Удаляет алиас по варианту написания."""
    global _dirty
    _ensure_loaded()
    aliases: dict = _data.get("product_aliases", {})
    if variant in aliases:
        del aliases[variant]
        _dirty = True
        _flush()


def resolve_product_name(name: str) -> str:
    """
    Возвращает каноническое название продукта.
    Если алиас не найден — возвращает исходное название.
    """
    _ensure_loaded()
    aliases: dict = _data.get("product_aliases", {})
    return aliases.get(name, name)


# ─────────────────────────── Шаблоны ──────────────────────────────────────

FIELD_LABELS: dict[str, str] = {
    "routeNumber": "№ маршрута",
    "address":     "Адрес",
    "product":     "Продукт",
    "unit":        "Ед. изм.",
    "quantity":    "Количество",
    "pcs":         "Шт",
    "productQty":  "Продукт (кол-во)",
}


def get_column_label(col: dict) -> str:
    """Returns display label for a column dict."""
    if col.get("label"):
        return col["label"]
    if col.get("merged") and col.get("productName"):
        return col["productName"]
    return FIELD_LABELS.get(col["field"], col["field"])


def create_template(name: str) -> dict:
    """Creates a new template with default columns and returns it."""
    global _dirty
    _ensure_loaded()
    tmpl = {
        "id": str(uuid.uuid4()),
        "name": name,
        "columns": [
            {"field": "routeNumber", "label": None, "merged": False},
            {"field": "address",     "label": None, "merged": False},
            {"field": "product",     "label": None, "merged": False},
            {"field": "quantity",    "label": None, "merged": False},
        ],
        "deptKey": None,
    }
    _data["templates"].append(tmpl)
    _dirty = True
    _flush()
    return tmpl


def delete_template(template_id: str) -> bool:
    """Deletes a template by id. Returns True if found and deleted."""
    global _dirty
    _ensure_loaded()
    templates: list = _data.get("templates", [])
    for i, t in enumerate(templates):
        if t["id"] == template_id:
            templates.pop(i)
            _dirty = True
            _flush()
            return True
    return False


def save_template(template_id: str, name: str, columns: list,
                  dept_key=None, fmt: str = "") -> bool:
    """Updates an existing template's name, columns, deptKey and format."""
    global _dirty
    _ensure_loaded()
    templates: list = _data.get("templates", [])
    for t in templates:
        if t["id"] == template_id:
            t["name"] = name
            t["columns"] = columns
            t["deptKey"] = dept_key
            t["format"] = fmt
            _dirty = True
            _flush()
            return True
    return False
