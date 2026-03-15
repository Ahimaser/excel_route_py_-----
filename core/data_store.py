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
import re
import shutil
import tempfile
import uuid
from pathlib import Path
from typing import Any, Iterable

# Регулярка для кода учреждения: только первые 3–4 цифры (109/1 → 109, 1391/2 → 1391)
_INSTITUTION_RE = re.compile(r"^\s*(\d{3,4})\b")

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
            "name": "Стандартный",
            "columns": [
                {"field": "routeNumber", "label": None, "merged": False},
                {"field": "address",     "label": None, "merged": False},
                {"field": "product",     "label": None, "merged": False},
                {"field": "unit",        "label": None, "merged": False},
                {"field": "quantity",    "label": None, "merged": False},
                {"field": "pcs",         "label": None, "merged": False},
            ],
        "deptKey": None,
        "deptKeys": [],
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
            "deptKeys": [],
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
        # Список кодов учреждений (3–4 цифры), для которых Шт округляются в большую сторону.
        "alwaysRoundUpInstitutions": [],
        # Адреса, для которых округление отключено (даже если учреждение в списке).
        "excludeRoundUpAddresses": [],
        # % от 1 шт по умолчанию (при остатке ≥ этого процента — округление в большую).
        "roundUpInstitutionPercent": 20,
        # % по отделам: {ключ отдела: процент}. Если отдел не указан — используется roundUpInstitutionPercent.
        "roundUpPercentByDept": {},
        # Печать этикеток: последний выбранный принтер.
        "labelsLastPrinter": "",
        # Отступы печати этикеток по умолчанию (см): верх/право=2, низ/лево=0.
        "labelsPrintMargins": {
            "top_cm": 2.0,
            "right_cm": 2.0,
            "bottom_cm": 0.0,
            "left_cm": 0.0,
        },
        # Очистка временных файлов после закрытия preview.
        "labelsTempAutoCleanup": True,
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
    global _data, _path, _dirty
    if _data is not None:
        return
    _path = _get_data_path()
    if _path.exists():
        try:
            with open(_path, "r", encoding="utf-8") as f:
                _data = json.load(f)
        except (json.JSONDecodeError, OSError) as e:
            log.warning("Не удалось загрузить store.json: %s", e)
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
            # Если columns пустой, но есть grid — строим columns из grid
            grid, merges = tmpl.get("grid"), tmpl.get("merges") or []
            if grid and merges is not None:
                cols_from_grid = _columns_from_grid(grid, merges)
                meaningful = [c for c in cols_from_grid if c.get("field") or ((c.get("label") or "").strip())]
                existing = tmpl.get("columns") or []
                existing_meaningful = len([c for c in existing if c.get("field") or ((c.get("label") or "").strip())])
                if not existing:
                    tmpl["columns"] = cols_from_grid
                    _dirty = True
                elif len(meaningful) > existing_meaningful:
                    tmpl["columns"] = cols_from_grid
                    _dirty = True
            # Миграция: лишние столбцы address → productQty (иначе дублируется адрес после ед. изм.)
            cols = tmpl.get("columns") or []
            address_count = sum(1 for c in cols if c.get("field") == "address")
            if address_count > 1:
                first = True
                for c in cols:
                    if c.get("field") == "address":
                        if first:
                            first = False
                        else:
                            c["field"] = "productQty"
                            c["label"] = None
                            c["merged"] = False
                            c.pop("productName", None)
                _dirty = True
            if "deptKeys" not in tmpl:
                dk = tmpl.get("deptKey")
                tmpl["deptKeys"] = [dk] if dk else []
            if "forGeneral" not in tmpl:
                tmpl["forGeneral"] = True
            if "forDepartments" not in tmpl:
                tmpl["forDepartments"] = True
    # Миграция: шаблон «Стандартный» — для отделов без явной привязки
    templates_list = _data.get("templates") or []
    has_standard = any((t.get("name") or "").strip() == "Стандартный" for t in templates_list)
    if not has_standard and templates_list:
        templates_list[0]["name"] = "Стандартный"
        _dirty = True
    if "product_aliases" not in _data:
        _data["product_aliases"] = {}
    if "last_main_routes" not in _data:
        _data["last_main_routes"] = None
    if "last_increase_routes" not in _data:
        _data["last_increase_routes"] = None
    if "routes_history" in _data and isinstance(_data.get("routes_history"), list) and _data["routes_history"]:
        _migrate_routes_history_to_hybrid()
    elif "routes_history" in _data:
        del _data["routes_history"]
    _data["settings"] = _data.get("settings") or {}
    # Гарантируем наличие ключей округления по учреждениям
    if "alwaysRoundUpInstitutions" not in _data["settings"]:
        _data["settings"]["alwaysRoundUpInstitutions"] = []
    else:
        # Миграция: 109/1 → 109, 1391/2 → 1391 (только первые 3–4 цифры)
        old = _data["settings"]["alwaysRoundUpInstitutions"]
        if old and isinstance(old, list):
            normalized = set()
            for code in old:
                if isinstance(code, str) and code.strip():
                    m = _INSTITUTION_RE.match(code.strip())
                    if m:
                        normalized.add(m.group(1))
                    else:
                        normalized.add(code.strip())
            new_list = sorted(normalized)
            if new_list != old:
                _data["settings"]["alwaysRoundUpInstitutions"] = new_list
                _dirty = True
                _flush()
    if "excludeRoundUpAddresses" not in _data["settings"]:
        _data["settings"]["excludeRoundUpAddresses"] = []
    if "roundUpInstitutionPercent" not in _data["settings"]:
        _data["settings"]["roundUpInstitutionPercent"] = 20
    if "roundUpPercentByDept" not in _data["settings"]:
        _data["settings"]["roundUpPercentByDept"] = {}
    if "labelsLastPrinter" not in _data["settings"]:
        _data["settings"]["labelsLastPrinter"] = ""
    if "labelsPrintMargins" not in _data["settings"] or not isinstance(_data["settings"].get("labelsPrintMargins"), dict):
        _data["settings"]["labelsPrintMargins"] = {
            "top_cm": 2.0,
            "right_cm": 2.0,
            "bottom_cm": 0.0,
            "left_cm": 0.0,
        }
    if "labelsTempAutoCleanup" not in _data["settings"]:
        _data["settings"]["labelsTempAutoCleanup"] = True
    if "productFileGroups" not in _data["settings"]:
        _data["settings"]["productFileGroups"] = {}
    # Миграция: переименование «ЧИЩЕНКА» → «Очищенные» (до проверки labelPrintMode) (до проверки labelPrintMode)
    for dept in _data.get("departments", []):
        for sub in dept.get("subdepts", []):
            sub_name = sub.get("name") or ""
            if sub_name.strip().upper() == "ЧИЩЕНКА":
                sub["name"] = "Очищенные"
                _dirty = True
    # Миграция: labelsFor, labelPrintMode, labelsEnabled для отделов/подотделов
    for dept in _data.get("departments", []):
        if dept.get("labelsFor") is None:
            dept["labelsFor"] = "both"
        if dept.get("labelPrintMode") is None:
            n = (dept.get("name") or "").lower()
            dept["labelPrintMode"] = (
                "chistchenka" if "очищенные" in n
                else "sypuchka" if "сыпучка" in n
                else "polufabricates" if "полуфаб" in n
                else "default"
            )
        if dept.get("labelsEnabled") is None:
            dept["labelsEnabled"] = True
        for sub in dept.get("subdepts", []):
            if sub.get("labelsFor") is None:
                sub["labelsFor"] = "both"
            if sub.get("labelPrintMode") is None:
                n = (sub.get("name") or "").lower()
                sub["labelPrintMode"] = (
                    "chistchenka" if "очищенные" in n
                    else "sypuchka" if "сыпучка" in n
                    else "polufabricates" if "полуфаб" in n
                    else "default"
                )
            if sub.get("labelsEnabled") is None:
                sub["labelsEnabled"] = True
    # Миграция: один отдел/подотдел — только один шаблон (убираем дубликаты deptKeys).
    # «Последний побеждает»: при конфликте dept_key остаётся у шаблона, идущего позже в списке.
    templates_list = _data.get("templates") or []
    seen_dept_keys: set[str] = set()
    for tmpl in reversed(templates_list):
        dk = tmpl.get("deptKeys") or []
        if not dk:
            continue
        new_dk = [k for k in dk if k not in seen_dept_keys]
        seen_dept_keys.update(dk)
        if new_dk != dk:
            tmpl["deptKeys"] = new_dk
            tmpl["deptKey"] = new_dk[0] if len(new_dk) == 1 else None
            _dirty = True
    if _dirty:
        _flush()


def _create_backup() -> None:
    """Создаёт резервные копии store.json (до 2 штук: .bak1, .bak2)."""
    if _path is None or not _path.exists():
        return
    try:
        bak2 = _path.with_suffix(".json.bak2")
        bak1 = _path.with_suffix(".json.bak1")
        if bak1.exists():
            shutil.copy2(bak1, bak2)
        shutil.copy2(_path, bak1)
    except OSError as e:
        log.warning("Не удалось создать резервную копию store.json: %s", e)


def _flush() -> None:
    """Атомарная запись данных на диск через временный файл."""
    global _dirty
    if not _dirty or _data is None or _path is None:
        return
    try:
        _create_backup()
        dir_ = _path.parent
        with tempfile.NamedTemporaryFile(
            "w", encoding="utf-8", dir=dir_, delete=False, suffix=".tmp"
        ) as tf:
            json.dump(_data, tf, ensure_ascii=False, indent=2)
            tmp_path = tf.name
        # Атомарная замена (работает на Windows и Linux)
        os.replace(tmp_path, _path)
        _dirty = False
    except (OSError, json.JSONEncodeError) as e:
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
    Конвертация «В Грязные» доступна только для подотдела Чищенка — при смене отдела сбрасывается.
    """
    global _dirty
    _ensure_loaded()
    products: list = _data.get("products", [])
    for p in products:
        if p.get("name") == name:
            p.update(kwargs)
            if "deptKey" in kwargs and not is_subdept_chistchenka(kwargs.get("deptKey")):
                if p.get("showInDirty"):
                    p["showInDirty"] = False
                    if p.get("quantityMultiplier") == 1.25:
                        p["quantityMultiplier"] = None
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


def add_product(name: str, unit: str = "", dept_key: str | None = None) -> bool:
    """
    Добавляет продукт в справочник вручную.
    Возвращает True, если добавлен; False, если продукт с таким именем уже есть.
    """
    global _dirty
    _ensure_loaded()
    name = (name or "").strip()
    if not name:
        return False
    products: list = _data.get("products", [])
    if any(p.get("name") == name for p in products):
        return False
    products.append({
        "name": name,
        "unit": (unit or "").strip(),
        "deptKey": dept_key if dept_key else None,
    })
    _dirty = True
    _flush()
    return True


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
GRID_COLS = 6  # уменьшено с 8 до 6

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


def get_template_columns(tmpl: dict) -> list[dict]:
    """
    Возвращает список столбцов шаблона. Если columns пустой, но есть grid — строит из grid.
    """
    cols = tmpl.get("columns", [])
    if cols:
        return cols
    grid = tmpl.get("grid")
    merges = tmpl.get("merges") or []
    if grid and len(grid) > 0:
        return _columns_from_grid(grid, merges)
    return []


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


def get_department_display_name(dept_key: str) -> str:
    """Возвращает отображаемое имя отдела/подотдела по ключу."""
    _ensure_loaded()
    for dept in _data.get("departments", []):
        if dept.get("key") == dept_key:
            return dept.get("name") or dept_key
        for sub in dept.get("subdepts", []):
            if sub.get("key") == dept_key:
                return sub.get("name") or dept_key
    return dept_key


def is_subdept_chistchenka(dept_key: str | None) -> bool:
    """
    True, если dept_key — подотдел «Очищенные» (по имени или labelPrintMode).
    Конвертация «В Грязные» доступна только для такого подотдела.
    """
    if not dept_key:
        return False
    _ensure_loaded()
    for dept in _data.get("departments", []):
        for sub in dept.get("subdepts", []):
            if sub.get("key") == dept_key:
                mode = sub.get("labelPrintMode")
                if mode == "chistchenka":
                    return True
                name = (sub.get("name") or "").lower()
                return "очищенные" in name
    return False


def is_subdept_polufabricates(dept_key: str | None) -> bool:
    """
    True, если dept_key — подотдел «Полуфабрикаты» (по имени или labelPrintMode).
    Используется для расчёта хвостиков (pcsTail) при отображении шт.
    """
    if not dept_key:
        return False
    _ensure_loaded()
    for dept in _data.get("departments", []):
        for sub in dept.get("subdepts", []):
            if sub.get("key") == dept_key:
                mode = sub.get("labelPrintMode")
                if mode in ("polufabricates", "polufabrikaty"):
                    return True
                name = (sub.get("name") or "").lower()
                return "полуфаб" in name
        # Проверяем и отдел (полуфабрикаты может быть отделом, а не подотделом)
        if dept.get("key") == dept_key:
            mode = dept.get("labelPrintMode")
            if mode in ("polufabricates", "polufabrikaty"):
                return True
            name = (dept.get("name") or "").lower()
            return "полуфаб" in name
    return False


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
    Строка 0 может быть заголовком (объединённая) — тогда берём строку 1 для столбцов.
    """
    if not grid or len(grid) == 0:
        return []
    num_cols = len(grid[0]) if grid[0] else 0
    if num_cols == 0:
        return []

    def _extract_cols_from_row(row_idx: int) -> list[dict]:
        out = []
        for c in range(num_cols):
            is_covered = False
            for (r0, c0, rs, cs) in merges:
                if r0 == row_idx and c0 < c < c0 + cs:
                    is_covered = True
                    break
            if is_covered:
                continue
            cell = grid[row_idx][c] if c < len(grid[row_idx]) else {"text": "", "field": None}
            col_span = 1
            for (r0, c0, rs, cs) in merges:
                if r0 == row_idx and c0 == c:
                    col_span = cs
                    break
            label = (cell.get("text") or "").strip() or None
            field = cell.get("field")
            if not field and label:
                for fk, fv in FIELD_LABELS.items():
                    if fv == label:
                        field = fk
                        break
            # Пустые ячейки — productQty без productName (записывается ""), не address (иначе дублируется адрес)
            col = {"field": field or "productQty", "label": label, "merged": col_span > 1}
            if col_span > 1 and label:
                col["productName"] = label
            out.append(col)
        return out

    # Строка 0 полностью объединена (заголовок) — столбцы в строке 1
    row0_merged = any(r0 == 0 and c0 == 0 and cs >= num_cols for (r0, c0, rs, cs) in merges)
    cols = _extract_cols_from_row(1) if row0_merged and len(grid) > 1 else _extract_cols_from_row(0)
    # Fallback: если строка 1 пустая (шаблон из columns), пробуем строку 0
    if len(cols) < 2 and row0_merged and len(grid) > 0:
        cols0 = _extract_cols_from_row(0)
        if len(cols0) > len(cols):
            cols = cols0
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
        "deptKeys": [],
        "forGeneral": True,
        "forDepartments": True,
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
    dept_keys: list | None = None,
    grid: list | None = None,
    merges: list | None = None,
    grid_rows: int | None = None,
    grid_cols: int | None = None,
    title_row: dict | None = None,
    for_general: bool = True,
    for_departments: bool = True,
) -> bool:
    """Обновляет шаблон: имя, столбцы, отделы, сетка, размер, заголовок, область применения.
    dept_keys: список ключей отделов/подотделов (пустой = шаблон по умолчанию).
    for_general: применять к общим маршрутам; for_departments: к маршрутам по отделам."""
    global _dirty
    _ensure_loaded()
    if dept_keys is None:
        dept_keys = [dept_key] if dept_key else []
    dept_keys = list(dept_keys)
    templates: list = _data.get("templates", [])
    for t in templates:
        if t["id"] == template_id:
            # Один отдел/подотдел — только один шаблон: убираем dept_keys из всех других шаблонов
            if dept_keys:
                for other in templates:
                    if other["id"] != template_id:
                        dk = other.get("deptKeys") or []
                        if dk:
                            other["deptKeys"] = [k for k in dk if k not in dept_keys]
                            other["deptKey"] = other["deptKeys"][0] if len(other["deptKeys"]) == 1 else None
                            _dirty = True
            t["name"] = name
            if grid is not None and merges is not None:
                t["grid"] = grid
                t["merges"] = merges
                cols_from_grid = _columns_from_grid(grid, merges)
                meaningful = [c for c in cols_from_grid if c.get("field") or ((c.get("label") or "").strip())]
                existing = t.get("columns") or []
                if len(meaningful) >= 2:
                    t["columns"] = cols_from_grid
                elif existing and len(existing) >= 2:
                    t["columns"] = existing
                else:
                    t["columns"] = cols_from_grid
            else:
                t["columns"] = columns
            t["deptKeys"] = list(dept_keys)
            t["deptKey"] = dept_keys[0] if len(dept_keys) == 1 else None  # обратная совместимость
            t["forGeneral"] = for_general
            t["forDepartments"] = for_departments
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


# ─────────────────────────── Список учреждений по маршрутам ─────────────────

def get_institution_key_from_address(address: str) -> str | None:
    """
    Ключ учреждения по адресу маршрута.
    Только первые 3–4 цифры: 109/1 → 109, 1391/2 → 1391. Все адреса с одинаковыми
    цифрами — подразделения одного учреждения. Если цифр нет — адрес целиком.
    """
    s = (address or "").strip()
    if not s:
        return None
    m = _INSTITUTION_RE.match(s)
    return m.group(1) if m else s


def get_institution_list_from_routes(routes: Iterable[dict]) -> list[str]:
    """
    Список уникальных учреждений по маршрутам для округления.
    Только первые 3–4 цифры: 109/1 и 109/2 → 109; 1391/1, 1391/2 → 1391.
    Иначе — адрес целиком.
    """
    found: set[str] = set()
    for r in routes or []:
        key = get_institution_key_from_address(r.get("address") or "")
        if key:
            found.add(key)
    return sorted(found)


def get_institution_addresses_map(routes: Iterable[dict]) -> dict[str, list[str]]:
    """
    Словарь {код учреждения: [список адресов]}.
    Адреса с одинаковыми первыми 3–4 цифрами группируются под одним учреждением.
    """
    result: dict[str, list[str]] = {}
    seen: dict[str, set[str]] = {}
    for r in routes or []:
        addr = (r.get("address") or "").strip()
        if not addr:
            continue
        key = get_institution_key_from_address(addr)
        if not key:
            continue
        if key not in seen:
            seen[key] = set()
            result[key] = []
        if addr not in seen[key]:
            seen[key].add(addr)
            result[key].append(addr)
    for key in result:
        result[key].sort()
    return result


def build_dept_groups_from_routes(routes: Iterable[dict]) -> list[dict]:
    """
    Строит список групп {key, name, routes} для каждого отдела/подотдела.
    Используется при генерации файлов по отделам с учётом замен продуктов.
    """
    routes = list(routes or [])
    depts = get_ref("departments") or []
    products = get_ref("products") or []

    prod_by_dept: dict[str, list[str]] = {}
    for p in products:
        k = p.get("deptKey")
        if k:
            prod_by_dept.setdefault(k, []).append(p["name"])

    def _aggregate_products(prods: list[dict]) -> list[dict]:
        by_name: dict[str, dict] = {}
        for p in prods:
            name = p.get("name", "")
            if not name:
                continue
            if name in by_name:
                agg = by_name[name]
                try:
                    agg["quantity"] = float(agg.get("quantity") or 0) + float(p.get("quantity") or 0)
                except (ValueError, TypeError):
                    pass
            else:
                by_name[name] = dict(p)
        return list(by_name.values())

    def _collect_routes(dept_key: str) -> list[dict]:
        prod_names = set(prod_by_dept.get(dept_key, []))
        if not prod_names:
            return []
        result = []
        for r in routes:
            if r.get("excluded"):
                continue
            dept_prods = [p for p in r.get("products", []) if p["name"] in prod_names]
            if dept_prods:
                result.append({
                    "routeNum": r.get("routeNum", ""),
                    "address": r.get("address", ""),
                    "routeCategory": r.get("routeCategory") or "ШК",
                    "products": _aggregate_products(dept_prods),
                })
        return result

    groups: list[dict] = []
    for dept in depts:
        for sub in dept.get("subdepts", []):
            sub_routes = _collect_routes(sub["key"])
            if sub_routes:
                groups.append({
                    "key": sub["key"],
                    "name": sub["name"],
                    "is_subdept": True,
                    "parent_dept_name": dept.get("name") or dept.get("key", ""),
                    "routes": sub_routes,
                })
        dept_routes = _collect_routes(dept["key"])
        if dept_routes:
            groups.append({
                "key": dept["key"],
                "name": dept["name"],
                "is_subdept": False,
                "routes": dept_routes,
            })
    return groups


def get_round_up_percent_for_dept(dept_key: str | None) -> float:
    """
    % от 1 шт для округления по учреждениям.
    Если для отдела задан % в roundUpPercentByDept — возвращает его, иначе roundUpInstitutionPercent, по умолчанию 20.
    """
    if not dept_key:
        pct = get_setting("roundUpInstitutionPercent")
        return float(pct) if pct is not None else 20.0
    by_dept = get_setting("roundUpPercentByDept") or {}
    if isinstance(by_dept, dict) and dept_key in by_dept:
        try:
            return float(by_dept[dept_key])
        except (ValueError, TypeError):
            pass
    pct = get_setting("roundUpInstitutionPercent")
    return float(pct) if pct is not None else 20.0


# ─────────────────────────── История маршрутов (гибрид: индекс + файлы) ────

_HISTORY_DIR = "history"
_INDEX_FILENAME = "index.json"


def _get_history_dir() -> Path:
    """Папка истории маршрутов."""
    d = get_app_data_dir() / _HISTORY_DIR
    d.mkdir(parents=True, exist_ok=True)
    return d


def _get_history_index_path() -> Path:
    return _get_history_dir() / _INDEX_FILENAME


def _current_month_str() -> str:
    from datetime import datetime
    return datetime.now().strftime("%Y-%m")


def _read_history_index() -> list[dict]:
    """Читает индекс истории. Пустой список при ошибке."""
    path = _get_history_index_path()
    if not path.exists():
        return []
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return list(data) if isinstance(data, list) else []
    except (json.JSONDecodeError, OSError) as e:
        log.warning("Ошибка чтения индекса истории: %s", e)
        return []


def _write_history_index(entries: list[dict]) -> None:
    path = _get_history_index_path()
    try:
        with tempfile.NamedTemporaryFile(
            "w", encoding="utf-8", dir=path.parent, delete=False, suffix=".tmp"
        ) as tf:
            json.dump(entries, tf, ensure_ascii=False, indent=2)
            tmp_path = tf.name
        os.replace(tmp_path, path)
    except (OSError, json.JSONEncodeError) as e:
        log.error("Ошибка записи индекса истории: %s", e)


def _migrate_routes_history_to_hybrid() -> None:
    """Миграция: переносит routes_history из store.json в папку history/ (только текущий месяц)."""
    old = _data.get("routes_history") or []
    if not isinstance(old, list) or not old:
        return
    current_month = _current_month_str()
    hist_dir = _get_history_dir()
    index_entries: list[dict] = []
    for idx, entry in enumerate(old):
        if not isinstance(entry, dict):
            continue
        ts_val = str(entry.get("timestamp") or "")
        if ts_val[:7] != current_month:
            continue
        ts = ts_val[:19].replace(":", "-").replace("T", "_")
        ft = str(entry.get("fileType") or "main")
        fname = f"entry_{idx:04d}_{ts}_{ft}.json"
        fpath = hist_dir / fname
        try:
            with open(fpath, "w", encoding="utf-8") as f:
                json.dump(entry, f, ensure_ascii=False, indent=2)
            routes = entry.get("filteredRoutes") or entry.get("routes") or []
            index_entries.append({
                "timestamp": entry.get("timestamp"),
                "fileType": ft,
                "filename": fname,
                "routeCategory": entry.get("routeCategory") or "ШК",
                "count": len(routes),
            })
        except Exception as e:
            log.warning("Миграция записи истории %s: %s", fname, e)
    if index_entries:
        _write_history_index(index_entries)
    del _data["routes_history"]
    global _dirty
    _dirty = True


# ─────────────────────────── Последние маршруты ───────────────────────────

def save_last_routes(
    file_type: str,
    routes: list,
    unique_products: list,
    filtered_routes: list,
    route_category: str | None = None,
    save_dir: str | None = None,
) -> None:
    """Сохраняет данные маршрутов как последние (main или increase). route_category: ШК или СД. save_dir — папка сохранения (для истории)."""
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
    if save_dir:
        blob["saveDir"] = save_dir
    key = "last_main_routes" if file_type == "main" else "last_increase_routes"
    _data[key] = blob
    _dirty = True
    _flush()

    # Гибрид: один файл на день на тип — обновление вместо дублирования
    now = datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    now_iso = now.isoformat()
    fname = f"entry_{date_str}_{file_type}.json"
    hist_dir = _get_history_dir()
    fpath = hist_dir / fname

    # Если файл за сегодня уже есть — берём createdAt, обновляем modifiedAt
    created_at = now_iso
    if fpath.exists():
        try:
            with open(fpath, "r", encoding="utf-8") as f:
                existing = json.load(f)
            created_at = existing.get("createdAt") or existing.get("timestamp") or now_iso
        except Exception:
            pass

    entry_full = {
        "fileType": file_type,
        "date": date_str,
        "createdAt": created_at,
        "modifiedAt": now_iso,
        "routes": copy.deepcopy(routes),
        "uniqueProducts": copy.deepcopy(unique_products),
        "filteredRoutes": copy.deepcopy(filtered_routes),
    }
    if route_category:
        entry_full["routeCategory"] = route_category
    if save_dir:
        entry_full["saveDir"] = save_dir
    try:
        with open(fpath, "w", encoding="utf-8") as f:
            json.dump(entry_full, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log.error("Ошибка записи истории: %s", e)
        return

    # Удаляем старые файлы за этот же день (entry_*_type.json с датой в имени)
    for f in hist_dir.glob("entry_*.json"):
        if f.name == _INDEX_FILENAME or f.name == fname:
            continue
        if date_str in f.name and f.name.endswith(f"_{file_type}.json"):
            try:
                f.unlink()
            except OSError:
                pass

    count = len(filtered_routes or routes or [])
    index_entry = {
        "date": date_str,
        "createdAt": created_at,
        "modifiedAt": now_iso,
        "fileType": file_type,
        "filename": fname,
        "routeCategory": route_category or "ШК",
        "count": count,
    }

    index_entries = _read_history_index()
    # Обновляем или добавляем запись за этот день и тип
    index_entries = [e for e in index_entries if not (
        (e.get("date") == date_str or (e.get("timestamp") or "")[:10] == date_str)
        and e.get("fileType") == file_type
    )]
    index_entries.append(index_entry)

    # Только текущий месяц
    current_month = _current_month_str()
    def _entry_month(e):
        d = e.get("date") or (e.get("timestamp") or "")[:10]
        return d[:7] if d else ""
    index_entries = [e for e in index_entries if _entry_month(e) == current_month]
    _write_history_index(index_entries)

    # Удаляем файлы записей из прошлых месяцев
    kept = {e.get("filename") for e in index_entries}
    for f in hist_dir.glob("entry_*.json"):
        if f.name != _INDEX_FILENAME and f.name not in kept:
            try:
                f.unlink()
            except OSError:
                pass


def get_last_routes(file_type: str) -> dict | None:
    """Возвращает последние сохранённые маршруты (main или increase) или None."""
    _ensure_loaded()
    key = "last_main_routes" if file_type == "main" else "last_increase_routes"
    return _data.get(key)


def clear_last_routes() -> None:
    """Очищает последние сохранённые маршруты и историю (основной и довоз)."""
    global _dirty
    _ensure_loaded()
    _data["last_main_routes"] = None
    _data["last_increase_routes"] = None
    _dirty = True
    _flush()

    hist_dir = _get_history_dir()
    if hist_dir.exists():
        try:
            for f in hist_dir.iterdir():
                f.unlink()
        except OSError as e:
            log.warning("Ошибка очистки истории: %s", e)


def get_routes_history(file_type: str | None = None) -> list[dict]:
    """
    Возвращает метаданные истории (последние изменения сверху). Только за текущий месяц.
    file_type: "main" | "increase" | None (все).
    Каждый элемент: {date, createdAt, modifiedAt, fileType, filename, routeCategory, count}.
    """
    _ensure_loaded()
    index_entries = _read_history_index()
    current_month = _current_month_str()

    def _entry_month(e):
        d = e.get("date") or (e.get("timestamp") or "")[:10]
        return d[:7] if d else ""

    index_entries = [e for e in index_entries if _entry_month(e) == current_month]
    if file_type in ("main", "increase"):
        index_entries = [e for e in index_entries if e.get("fileType") == file_type]
    # Сортировка по modifiedAt (или timestamp) — последние сверху
    def _sort_key(e):
        m = e.get("modifiedAt") or e.get("timestamp") or ""
        return m
    index_entries.sort(key=_sort_key, reverse=True)
    return index_entries


def load_routes_history_entry(filename: str) -> dict | None:
    """Загружает полную запись истории по имени файла. Нормализует старый формат (timestamp → createdAt/modifiedAt)."""
    hist_dir = _get_history_dir()
    fpath = hist_dir / filename
    if not fpath.exists() or fpath.name == _INDEX_FILENAME:
        return None
    try:
        with open(fpath, "r", encoding="utf-8") as f:
            data = json.load(f)
        # Нормализация старого формата (только timestamp)
        if "createdAt" not in data and data.get("timestamp"):
            data["createdAt"] = data["timestamp"]
            data["modifiedAt"] = data["timestamp"]
            data["date"] = (data["timestamp"] or "")[:10]
        return data
    except (json.JSONDecodeError, OSError) as e:
        log.warning("Ошибка загрузки записи истории %s: %s", filename, e)
        return None


def list_backups() -> list[tuple[int, str, float]]:
    """
    Возвращает список доступных резервных копий store.json.
    Каждый элемент: (индекс 0/1, имя файла, mtime).
    """
    _ensure_loaded()
    if _path is None:
        return []
    result: list[tuple[int, str, float]] = []
    for i, suf in enumerate([".json.bak1", ".json.bak2"]):
        p = _path.with_suffix(suf)
        if p.exists():
            try:
                mtime = p.stat().st_mtime
                result.append((i, p.name, mtime))
            except OSError:
                pass
    return result


def restore_from_backup(backup_index: int) -> bool:
    """
    Восстанавливает store.json из резервной копии (0 = bak1, 1 = bak2).
    Возвращает True при успехе. После восстановления нужно перезапустить приложение.
    """
    _ensure_loaded()
    if _path is None:
        return False
    suf = ".json.bak1" if backup_index == 0 else ".json.bak2"
    src = _path.with_suffix(suf)
    if not src.exists():
        return False
    try:
        global _data, _dirty
        with open(src, "r", encoding="utf-8") as f:
            _data = json.load(f)
        shutil.copy2(src, _path)
        _dirty = False
        return True
    except (json.JSONDecodeError, OSError) as e:
        log.error("Ошибка восстановления из резервной копии: %s", e)
        return False


def delete_routes_history_entry(filename: str) -> bool:
    """Удаляет запись истории по имени файла. Возвращает True при успехе."""
    hist_dir = _get_history_dir()
    index_entries = _read_history_index()
    index_entries = [e for e in index_entries if e.get("filename") != filename]
    _write_history_index(index_entries)
    fpath = hist_dir / filename
    if fpath.exists() and fpath.name != _INDEX_FILENAME:
        try:
            fpath.unlink()
            return True
        except OSError as e:
            log.warning("Ошибка удаления файла истории %s: %s", filename, e)
    return False
