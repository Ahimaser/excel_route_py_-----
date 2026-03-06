"""
data_store.py — JSON хранилище настроек приложения.

Все данные (отделы, продукты, шаблоны, настройки, последние маршруты и т.д.)
сохраняются в store.json и подгружаются при следующем запуске программы.

Оптимизации:
- Прямой доступ к данным без лишних deep-copy (get_ref/set_key)
- Атомарная запись через временный файл (нет риска повреждения)
- Ленивая загрузка (load при первом обращении)
- Отложенная запись через _dirty-флаг (flush только при реальных изменениях)
- Кэш пути к рабочему столу
"""
from __future__ import annotations

import copy
import json
import logging
import os
import tempfile
import uuid
from pathlib import Path
from typing import Any

# Имя папки данных (не менять — сохраняет совместимость с существующими данными)
APP_NAME = "ExcelRouteManager"
log = logging.getLogger(__name__)

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
        "defaultFontSize": 12,
        "defaultMarginTop": 1.5,
        "defaultMarginLeft": 1.5,
        "defaultMarginBottom": 0.5,
        "defaultMarginRight": 0.5,
    },
    "last_main_routes": None,
    "last_increase_routes": None,
}

# ─────────────────────────── Состояние модуля ─────────────────────────────

_data: dict[str, Any] | None = None
_path: Path | None = None
_dirty: bool = False
_desktop_cache: str | None = None


# ─────────────────────────── Внутренние утилиты ───────────────────────────

def get_app_data_dir() -> Path:
    """Возвращает папку данных приложения (APPDATA/ExcelRouteManager или ~/.config/ExcelRouteManager)."""
    if os.name == "nt":
        base = Path(os.environ.get("APPDATA", Path.home()))
    else:
        base = Path.home() / ".config"
    folder = base / APP_NAME
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def _get_data_path() -> Path:
    return get_app_data_dir() / "store.json"


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
    if "last_main_routes" not in _data:
        _data["last_main_routes"] = None
    if "last_increase_routes" not in _data:
        _data["last_increase_routes"] = None
    _data["settings"] = _data.get("settings") or {}
    # Миграция: labelsFor, labelPrintMode, labelsEnabled для отделов/подотделов
    for dept in _data.get("departments", []):
        if dept.get("labelsFor") is None:
            dept["labelsFor"] = "both"
        if dept.get("labelPrintMode") is None:
            n = (dept.get("name") or "").lower()
            dept["labelPrintMode"] = "chistchenka" if "чищенка" in n else "sypuchka" if "сыпучка" in n else "default"
        if dept.get("labelsEnabled") is None:
            dept["labelsEnabled"] = True
        for sub in dept.get("subdepts", []):
            if sub.get("labelsFor") is None:
                sub["labelsFor"] = "both"
            if sub.get("labelPrintMode") is None:
                n = (sub.get("name") or "").lower()
                sub["labelPrintMode"] = "chistchenka" if "чищенка" in n else "sypuchka" if "сыпучка" in n else "default"
            if sub.get("labelsEnabled") is None:
                sub["labelsEnabled"] = True


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
        log.error("Ошибка записи store: %s", e)


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


def set_product_label_template(name: str, template_path: str) -> bool:
    """
    Устанавливает путь к шаблону этикетки для продукта.
    Возвращает True, если продукт найден и обновлён.
    """
    global _dirty
    _ensure_loaded()
    name = (name or "").strip()
    if not name:
        return False
    products: list = _data.get("products", [])
    for p in products:
        if p.get("name") == name:
            p["labelTemplatePath"] = template_path
            _dirty = True
            _flush()
            return True
    return False


def get_setting(key: str) -> Any:
    """Возвращает значение настройки (например defaultSaveDir, showPcsInPreview)."""
    _ensure_loaded()
    settings = _data.get("settings") or {}
    return settings.get(key)


def set_setting(key: str, value: Any) -> None:
    """Устанавливает одну настройку, не затирая остальные."""
    global _dirty
    _ensure_loaded()
    settings = dict(_data.get("settings") or {})
    settings[key] = value
    _data["settings"] = settings
    _dirty = True
    _flush()


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
    Продукты без имени пропускаются.
    """
    _ensure_loaded()
    products = _data.get("products", [])
    return {p["name"]: p for p in products if p.get("name")}


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


def remove_product(name: str) -> bool:
    """
    Удаляет продукт из справочника по имени и все связанные алиасы
    (где продукт — вариант или каноническое название).
    Возвращает True, если продукт был найден и удалён.
    """
    global _dirty
    _ensure_loaded()
    products: list = _data.get("products", [])
    new_products = [p for p in products if p.get("name") != name]
    if len(new_products) == len(products):
        return False
    aliases: dict = _data.get("product_aliases", {})
    to_remove = [v for v, c in aliases.items() if v == name or c == name]
    for v in to_remove:
        del aliases[v]
    _data["products"] = new_products
    _dirty = True
    _flush()
    return True


def resolve_product_name(name: str) -> str:
    """
    Возвращает каноническое название продукта.
    Если алиас не найден — возвращает исходное название.
    """
    _ensure_loaded()
    aliases: dict = _data.get("product_aliases", {})
    return aliases.get(name, name)


# ─────────────────────────── Шаблоны ──────────────────────────────────────

# Сетка редактора шаблона: 6 строк (3 — заголовки, 3 — данные), 8 столбцов
GRID_ROWS = 6
GRID_COLS = 8

FIELD_LABELS: dict[str, str] = {
    "routeNumber":  "№ маршрута",
    "address":      "Адрес",
    "product":      "Продукт",
    "unit":         "Ед. изм.",
    "quantity":     "Количество",
    "pcs":          "Шт",
    "productQty":   "Продукт (кол-во)",
    "productsWide": "Продукт (колонка на каждый)",
    "nomenclature": "Номенклатура",
}


def get_department_choices() -> list[tuple[str, str]]:
    """Список (key, display_name) для комбобокса привязки шаблона к отделу. Первый элемент — «Все отделы»."""
    _ensure_loaded()
    result = [("", "Все отделы")]
    for dept in _data.get("departments", []):
        key = dept.get("key") or ""
        name = dept.get("name") or key
        if key:
            result.append((key, name))
        for sub in dept.get("subdepts", []):
            sk = sub.get("key") or ""
            sn = sub.get("name") or sk
            if sk:
                result.append((sk, f"  └ {sn}"))
    return result


def get_column_label(col: dict) -> str:
    """Returns display label for a column dict."""
    if col.get("label"):
        return col["label"]
    if col.get("merged") and col.get("productName"):
        return col["productName"]
    return FIELD_LABELS.get(col["field"], col["field"])


def _default_grid() -> list:
    """Пустая сетка GRID_ROWS×GRID_COLS: каждая ячейка {text, field}."""
    return [
        [{"text": "", "field": None} for _ in range(GRID_COLS)]
        for _ in range(GRID_ROWS)
    ]


def _columns_from_grid(grid: list, merges: list) -> list:
    """
    Строит список столбцов (для экспорта) из сетки и объединений.
    merges: список (r, c, rowSpan, colSpan) — верхний левый угол и размер.
    """
    if not grid or len(grid) == 0:
        return []
    num_cols = len(grid[0]) if grid[0] else 0
    cols = []
    for c in range(num_cols):
        # Пропускаем ячейки, входящие в объединение слева
        is_covered = False
        for (r0, c0, rs, cs) in merges:
            if r0 == 0 and c0 < c < c0 + cs:
                is_covered = True
                break
        if is_covered:
            continue
        cell = grid[0][c] if c < len(grid[0]) else {"text": "", "field": None}
        row_span, col_span = 1, 1
        for (r0, c0, rs, cs) in merges:
            if r0 == 0 and c0 == c:
                row_span, col_span = rs, cs
                break
        label = (cell.get("text") or "").strip() or None
        field = cell.get("field")
        if not field and label:
            # Пытаемся сопоставить с известным полем по подписи
            for fk, fv in FIELD_LABELS.items():
                if fv == label:
                    field = fk
                    break
        col = {"field": field or "address", "label": label, "merged": col_span > 1}
        if col_span > 1 and label:
            col["productName"] = label
        cols.append(col)
    return cols if cols else [
        {"field": "routeNumber", "label": None, "merged": False},
        {"field": "address", "label": None, "merged": False},
    ]


def create_template(name: str) -> dict:
    """Создаёт новый шаблон с пустой сеткой 6×8 и возвращает его."""
    global _dirty
    _ensure_loaded()
    from datetime import date, timedelta
    tomorrow = (date.today() + timedelta(days=1)).strftime("%d.%m.%Y")
    tmpl = {
        "id": str(uuid.uuid4()),
        "name": name,
        "columns": [
            {"field": "routeNumber", "label": None, "merged": False},
            {"field": "address", "label": None, "merged": False},
            {"field": "product", "label": None, "merged": False},
            {"field": "quantity", "label": None, "merged": False},
        ],
        "deptKey": None,
        "format": "",
        "grid": _default_grid(),
        "merges": [],
        "gridRows": GRID_ROWS,
        "gridCols": GRID_COLS,
        "titleRow": {
            "auto": True,
            "includeDept": True,
            "date": tomorrow,
            "type": "main",
        },
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


def save_template(
    template_id: str,
    name: str,
    columns: list,
    dept_key=None,
    fmt: str = "",
    grid: list | None = None,
    merges: list | None = None,
    grid_rows: int | None = None,
    grid_cols: int | None = None,
    title_row: dict | None = None,
) -> bool:
    """Обновляет шаблон: имя, столбцы, отдел, формат, сетка, размер, заголовок."""
    global _dirty
    _ensure_loaded()
    templates: list = _data.get("templates", [])
    for t in templates:
        if t["id"] == template_id:
            t["name"] = name
            if grid is not None and merges is not None:
                t["grid"] = grid
                t["merges"] = merges
                t["columns"] = _columns_from_grid(grid, merges)
            else:
                t["columns"] = columns
            t["deptKey"] = dept_key
            t["format"] = fmt
            if grid_rows is not None:
                t["gridRows"] = grid_rows
            if grid_cols is not None:
                t["gridCols"] = grid_cols
            if title_row is not None:
                t["titleRow"] = title_row
            _dirty = True
            _flush()
            return True
    return False


# ─────────────────────────── Последние маршруты ───────────────────────────

def save_last_routes(
    file_type: str,
    routes: list,
    unique_products: list,
    filtered_routes: list,
    route_category: str | None = None,
) -> None:
    """Сохраняет данные маршрутов как последние (main или increase). route_category: ШК или СД."""
    global _dirty
    _ensure_loaded()
    from datetime import datetime
    blob = {
        "timestamp": datetime.now().isoformat(),
        "routes": copy.deepcopy(routes),
        "uniqueProducts": copy.deepcopy(unique_products),
        "filteredRoutes": copy.deepcopy(filtered_routes),
    }
    if route_category:
        blob["routeCategory"] = route_category
    key = "last_main_routes" if file_type == "main" else "last_increase_routes"
    _data[key] = blob
    _dirty = True
    _flush()


def get_last_routes(file_type: str) -> dict | None:
    """Возвращает последние сохранённые маршруты (main или increase) или None."""
    _ensure_loaded()
    key = "last_main_routes" if file_type == "main" else "last_increase_routes"
    return _data.get(key)


def clear_last_routes() -> None:
    """Очищает последние сохранённые маршруты (основной и довоз)."""
    global _dirty
    _ensure_loaded()
    _data["last_main_routes"] = None
    _data["last_increase_routes"] = None
    _dirty = True
    _flush()
