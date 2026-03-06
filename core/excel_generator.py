"""
excel_generator.py — Генератор XLS файлов маршрутов.

Форматы файлов по отделам:
  "wide"  (Формат 1) — каждый продукт в отдельном столбце:
           Маршрут | Адрес | ПродуктA | ПродуктB | ...
           Значение ячейки: "5 кг / 3 шт" (шт — опционально)
  "rows"  (Формат 2) — строчный:
           Строка маршрута: Маршрут | Адрес | — | —
           Строка продукта: — | Название продукта | Кол-во | Шт

Сортировка: по убыванию номера маршрута (числовая).
"""
from __future__ import annotations

import logging
import math
import os
import re
from datetime import date, timedelta
from typing import Any

import xlwt

from core import data_store

log = logging.getLogger("excel_generator")

ROUTE_SIGN = "\u2116"

# ─────────────────────────── Кэш стилей ──────────────────────────────────

_STYLES: dict[str, xlwt.XFStyle] | None = None
_STYLES_FONT_PT: int | None = None


def _get_styles() -> dict[str, xlwt.XFStyle]:
    """Возвращает набор стилей. Размер шрифта из настроек (по умолчанию 12pt). Кэш по font_pt."""
    global _STYLES, _STYLES_FONT_PT
    font_pt = 12
    try:
        v = data_store.get_setting("defaultFontSize")
        if v is not None:
            font_pt = max(8, min(24, int(v)))
    except Exception:
        pass
    if _STYLES is not None and _STYLES_FONT_PT == font_pt:
        return _STYLES
    height = font_pt * 20  # xlwt: height в 1/20 пункта

    font_bold = xlwt.Font()
    font_bold.bold = True
    font_bold.height = height

    font_normal = xlwt.Font()
    font_normal.height = height

    align_wrap = xlwt.Alignment()
    align_wrap.wrap = xlwt.Alignment.WRAP_AT_RIGHT
    align_wrap.vert = xlwt.Alignment.VERT_TOP

    align_center = xlwt.Alignment()
    align_center.horz = xlwt.Alignment.HORZ_CENTER
    align_center.vert = xlwt.Alignment.VERT_TOP

    align_top = xlwt.Alignment()
    align_top.vert = xlwt.Alignment.VERT_TOP

    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN

    def _make(font, alignment, has_borders=True):
        s = xlwt.XFStyle()
        s.font = font
        s.alignment = alignment
        if has_borders:
            s.borders = borders
        return s

    # Жёлтый фон для номера маршрута в этикетках
    pattern_yellow = xlwt.Pattern()
    pattern_yellow.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern_yellow.pattern_fore_colour = xlwt.Style.colour_map["yellow"]
    style_yellow = xlwt.XFStyle()
    style_yellow.font = font_normal
    style_yellow.alignment = align_top
    style_yellow.borders = borders
    style_yellow.pattern = pattern_yellow

    _STYLES = {
        "header":      _make(font_bold,   align_center),
        "header_wrap": _make(font_bold,   align_wrap),
        "cell":        _make(font_normal, align_top),
        "cell_wrap":   _make(font_normal, align_wrap),
        "num":         _make(font_normal, align_top),
        "title":       _make(font_bold,   align_top, has_borders=False),
        "cell_yellow": style_yellow,
    }
    _STYLES_FONT_PT = font_pt
    return _STYLES


def _apply_page_margins(ws: xlwt.Worksheet, for_labels: bool = False) -> None:
    """Применяет отступы страницы из настроек (см). Не применяется к этикеткам."""
    if for_labels:
        return
    try:
        top = float(data_store.get_setting("defaultMarginTop") or 1.5)
        left = float(data_store.get_setting("defaultMarginLeft") or 1.5)
        bottom = float(data_store.get_setting("defaultMarginBottom") or 0.5)
        right = float(data_store.get_setting("defaultMarginRight") or 0.5)
    except (TypeError, ValueError):
        top, left, bottom, right = 1.5, 1.5, 0.5, 0.5
    inch = 1.0 / 2.54
    ws.set_top_margin(top * inch)
    ws.set_left_margin(left * inch)
    ws.set_bottom_margin(bottom * inch)
    ws.set_right_margin(right * inch)


# ─────────────────────────── Утилиты ──────────────────────────────────────

def _tomorrow() -> date:
    return date.today() + timedelta(days=1)


def _format_date(d: date) -> str:
    return d.strftime("%d.%m.%Y")


def _type_label(file_type: str) -> str:
    return "Увеличение (Довоз)" if file_type == "increase" else "Основной"


def _type_suffix(file_type: str) -> str:
    return "УВ" if file_type == "increase" else "ОСН"


def calc_pcs(quantity: float, pcs_per_unit: float, round_up: bool = True) -> int:
    """
    Рассчитывает количество в штуках (округление в большую сторону при round_up=True).

    Для кг/л дополнительная логика в _apply_pcs:
      < 0.2 → 0 шт; < 1 → 0 шт; >= 1 → округление вверх.
    """
    if pcs_per_unit <= 0:
        return 0
    whole = math.floor(quantity / pcs_per_unit)
    remainder = quantity - whole * pcs_per_unit
    half = pcs_per_unit / 2
    return whole + (1 if (remainder >= half if round_up else remainder > half) else 0)


def _apply_pcs(routes: list[dict], prod_map: dict[str, dict]) -> list[dict]:
    """
    Добавляет к продуктам маршрутов displayQuantity (с учётом множителя замены) и pcs.
    prod_map: {name: product_settings_dict}. Коэффициент замены (quantityMultiplier), напр. 1.25
    для пересчёта очищенных → грязные: отображаемое количество = количество × коэффициент.
    Округление берётся по категории маршрута (ШК/СД).
    """
    for route in routes:
        route_cat = route.get("routeCategory") or "ШК"
        for prod in route.get("products", []):
            sp = prod_map.get(prod["name"], {})
            qty = prod.get("quantity")
            mult = float(sp.get("quantityMultiplier", 1.0) or 1.0)
            if qty is not None:
                try:
                    display_qty = float(qty) * mult
                    prod["displayQuantity"] = display_qty
                except (ValueError, TypeError):
                    prod["displayQuantity"] = qty
            else:
                prod["displayQuantity"] = qty

            pcs = None
            if sp.get("showPcs") and prod.get("unit", "").lower() != "шт":
                eff_qty = prod.get("displayQuantity")
                if eff_qty is not None:
                    try:
                        val = float(eff_qty)
                        unit_lower = (prod.get("unit") or "").strip().lower()
                        # Для кг/л: < 0.2 → 0 шт; < 1 → 0 шт; >= 1 → округление в большую сторону
                        if unit_lower in ("кг", "л", "kg", "l"):
                            if val < 0.2:
                                pcs = 0
                            elif val < 1:
                                pcs = 0
                            else:
                                if route_cat == "СД":
                                    round_up = sp.get("roundUpСД") if "roundUpСД" in sp else sp.get("roundUp", True)
                                else:
                                    round_up = sp.get("roundUpШК") if "roundUpШК" in sp else sp.get("roundUp", True)
                                pcs = calc_pcs(val, float(sp.get("pcsPerUnit", 1)), bool(round_up))
                        else:
                            if route_cat == "СД":
                                round_up = sp.get("roundUpСД") if "roundUpСД" in sp else sp.get("roundUp", True)
                            else:
                                round_up = sp.get("roundUpШК") if "roundUpШК" in sp else sp.get("roundUp", True)
                            pcs = calc_pcs(val, float(sp.get("pcsPerUnit", 1)), bool(round_up))
                    except (ValueError, TypeError):
                        pass
            prod["pcs"] = pcs
    return routes


# Символы, недопустимые в имени листа Excel: \ / ? * [ ]
_SHEET_NAME_FORBIDDEN = re.compile(r'[\\/?*\[\]]')


def _safe_sheet_name(name: str) -> str:
    """Имя листа Excel: макс 31 символ, без \\ / ? * [ ]."""
    s = str(name).strip()
    s = _SHEET_NAME_FORBIDDEN.sub("_", s)
    s = s.strip("_") or "Лист"
    return s[:31]


def _unique_sheet_name(name: str, used: set[str]) -> str:
    """Генерирует уникальное имя листа (макс 31 символ, без недопустимых символов)."""
    base = _safe_sheet_name(name)
    if len(base) > 28:
        base = base[:28]
    candidate = base
    counter = 2
    while candidate in used:
        candidate = f"{base[:24]}_{counter}"
        counter += 1
    used.add(candidate)
    return candidate


def _set_col_width(sheet: xlwt.Worksheet, col: int, width_chars: int) -> None:
    sheet.col(col).width = min(width_chars * 256, 65535)


def _sort_routes(routes: list[dict], sort_asc: bool = False) -> list[dict]:
    """Сортирует маршруты по номеру маршрута (числовая сортировка).
    sort_asc=False (по умолчанию) — по убыванию.
    sort_asc=True — по возрастанию.
    Маршруты с неопределённым номером — всегда в начало.
    """
    def _key(r: dict):
        num = r.get("routeNum", "")
        try:
            n = int(str(num).strip())
            return (1, n if sort_asc else -n)
        except (ValueError, TypeError):
            return (0, 0)
    return sorted(routes, key=_key)


def _safe_filename(name: str) -> str:
    """Убирает запрещённые символы из имени файла."""
    return re.sub(r'[\\/:*?"<>|]', "_", name)


def extract_house_number(address: str) -> str:
    """
    Извлекает из адреса номер дома, строения, корпуса и/или цифры перед № (U+2116).
    Поддерживает смешанный формат: "дом 3 строение 2", "д. 2 корпус 1", "д.5", "стр. 2" и т.д.
    Возвращает одну строку с найденными значениями через запятую (например "3, стр. 2, корп. 1").
    """
    if not address:
        return ""
    s = str(address).strip()
    parts: list[str] = []

    # дом / д. — дом 3, д. 5, д.5а, д. 109/1
    m = re.search(r"(?:^|[^\w])(?:дом\s*|д\.\s*)(\d+(?:/\d+)?[а-яА-Яa-zA-Z]*)", s, re.IGNORECASE)
    if m:
        parts.append(m.group(1).strip())

    # строение / строен. / стр.
    m = re.search(r"(?:строение|строен\.?|стр\.)\s*(\d+)", s, re.IGNORECASE)
    if m:
        parts.append(f"стр. {m.group(1)}")

    # корпус / корп.
    m = re.search(r"корпус?\s*(\d+)", s, re.IGNORECASE)
    if m:
        parts.append(f"корп. {m.group(1)}")

    # цифры перед символом № (U+2116)
    m = re.search(r"(\d+)\s*[№\u2116]", s)
    if m:
        num = m.group(1)
        if num not in parts:  # не дублировать, если уже как "дом"
            parts.append(num)

    return ", ".join(parts) if parts else ""


def _label_print_mode_for_dept(dept_key: str | None, departments_ref: list | None) -> str:
    """Режим печати этикеток: default, chistchenka, sypuchka (по имени или labelPrintMode)."""
    if not dept_key or not departments_ref:
        return "default"
    for dept in departments_ref:
        if dept.get("key") == dept_key:
            mode = dept.get("labelPrintMode")
            if mode in ("chistchenka", "sypuchka"):
                return mode
            name = (dept.get("name") or "").lower()
            if "чищенка" in name:
                return "chistchenka"
            if "сыпучка" in name:
                return "sypuchka"
            return "default"
        for sub in dept.get("subdepts", []):
            if sub.get("key") == dept_key:
                mode = sub.get("labelPrintMode")
                if mode in ("chistchenka", "sypuchka"):
                    return mode
                name = (sub.get("name") or "").lower()
                if "чищенка" in name:
                    return "chistchenka"
                if "сыпучка" in name:
                    return "sypuchka"
                return "default"
    return "default"


def _label_rules_for_dept(dept_key: str | None, departments_ref: list | None) -> dict:
    """
    Возвращает правила этикеток для отдела/подотдела: labelRules или пустой dict.
    labelRules может содержать chistchenka: {maxKgPerLabel} (макс. кг на этикетку; дубликаты по 5 кг, остаток — последняя этикетка),
    sypuchka: {thresholdKg, labelAbove, labelBelow}.
    """
    if not dept_key or not departments_ref:
        return {}
    for dept in departments_ref:
        if dept.get("key") == dept_key:
            return dict(dept.get("labelRules") or {})
        for sub in dept.get("subdepts", []):
            if sub.get("key") == dept_key:
                return dict(sub.get("labelRules") or {})
    return {}


def _dept_display_name(dept_key: str | None, departments_ref: list | None) -> str:
    """Возвращает отображаемое имя отдела/подотдела по key."""
    if not dept_key or not departments_ref:
        return ""
    for dept in departments_ref:
        if dept.get("key") == dept_key:
            return dept.get("name") or dept_key
        for sub in dept.get("subdepts", []):
            if sub.get("key") == dept_key:
                parent = dept.get("name") or ""
                sub_name = sub.get("name") or dept_key
                return f"{parent} / {sub_name}" if parent else sub_name
    return dept_key


def _route_sort_key_labels(route_num: str) -> tuple[int, int, str]:
    """Ключ сортировки номера маршрута по возрастанию (для этикеток: сначала по адресу, затем по этому ключу)."""
    try:
        return (1, int(str(route_num).strip()), str(route_num))
    except (ValueError, TypeError):
        return (0, 0, str(route_num))


def _label_sort_key_route(route: dict) -> tuple[str, int, int, str]:
    """Ключ сортировки для этикеток: сначала адрес, затем номер маршрута по возрастанию."""
    address = (route.get("address") or "").strip()
    route_num = str(route.get("routeNum", ""))
    return (address, *_route_sort_key_labels(route_num))


def _load_template_matrix(template_path: str) -> tuple[int, int, list, int]:
    """
    Загружает шаблон XLS в матрицу (nrows, ncols, list of rows, last_filled_row).
    last_filled_row — индекс последней строки, в которой есть хотя бы одна непустая ячейка (-1 если пусто).
    После этой строки программа добавляет одну строку с данными (№ маршрута, дом/строение, количество).
    """
    import xlrd
    wb = xlrd.open_workbook(template_path, formatting_info=False)
    sheet = wb.sheet_by_index(0)
    nrows, ncols = sheet.nrows, sheet.ncols
    matrix = []
    last_filled = -1
    for r in range(nrows):
        row = []
        for c in range(ncols):
            cell = sheet.cell(r, c)
            if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                row.append("")
            elif cell.ctype == xlrd.XL_CELL_NUMBER:
                v = cell.value
                row.append(int(v) if v == int(v) else v)
            else:
                row.append(str(cell.value).strip() if cell.value else "")
        matrix.append(row)
        if any(row):
            last_filled = r
    return nrows, ncols, matrix, last_filled


def load_label_template_matrix(template_path: str) -> tuple[int, int, list, int]:
    """
    Публичная обёртка для загрузки шаблона этикетки (предпросмотр в UI).
    Возвращает (nrows, ncols, matrix, last_filled_row).
    """
    return _load_template_matrix(template_path)


def _write_label_block(
    ws: Any,
    matrix: list,
    template_rows: int,
    ncols: int,
    start_row: int,
    route_num: str,
    house: str,
    qty_val: Any,
    styles: dict,
    label_layout: list[dict] | None = None,
) -> None:
    """
    Пишет один блок этикетки: копия первых template_rows строк шаблона (matrix),
    затем заполнение ячеек по данным. Если задан label_layout — список
    {"row": r, "col": c, "field": "routeNumber"|"house"|"quantity"}, значения пишутся
    в указанные ячейки (start_row + r, c). Иначе по умолчанию: столбцы 0, 1, 2 в строке данных.
    """
    style_yellow = styles.get("cell_yellow", styles["cell"])
    style_num = styles.get("num", styles["cell"])
    # Копируем строки шаблона (0 .. template_rows-1)
    for r in range(template_rows):
        for c in range(ncols):
            val = matrix[r][c] if r < len(matrix) and c < len(matrix[r]) else ""
            cell_style = style_num if isinstance(val, (int, float)) else styles["cell"]
            try:
                if isinstance(val, (int, float)):
                    ws.write(start_row + r, c, val, cell_style)
                else:
                    ws.write(start_row + r, c, str(val), cell_style)
            except Exception as _e:
                log.debug("_write_label_block: row=%d col=%d err=%s", start_row + r, c, _e)
                ws.write(start_row + r, c, str(val), cell_style)
    # Заполнение данных: по layout или по умолчанию (строка данных, колонки 0,1,2)
    data_row = start_row + template_rows
    ncols_write = max(ncols, 3)
    if label_layout:
        values_by_cell: dict[tuple[int, int], Any] = {}
        for pl in label_layout:
            r, c = pl.get("row", template_rows), pl.get("col", 0)
            f = pl.get("field")
            if f == "routeNumber":
                values_by_cell[(r, c)] = (route_num, style_yellow)
            elif f == "house":
                values_by_cell[(r, c)] = (house, styles["cell"])
            elif f == "quantity":
                values_by_cell[(r, c)] = (qty_val if qty_val is not None else "", style_num)
        for (r, c), (val, cell_style) in values_by_cell.items():
            try:
                row_abs = start_row + r
                if isinstance(val, (int, float)):
                    ws.write(row_abs, c, val, cell_style)
                else:
                    ws.write(row_abs, c, str(val), cell_style)
            except Exception as _e:
                log.debug("_write_label_block layout row=%d col=%d err=%s", start_row + r, c, _e)
    else:
        for c in range(ncols_write):
            if c == 0:
                val, cell_style = route_num, style_yellow
            elif c == 1:
                val, cell_style = house, styles["cell"]
            elif c == 2:
                val, cell_style = (qty_val if qty_val is not None else ""), style_num
            else:
                val, cell_style = "", styles["cell"]
            try:
                if isinstance(val, (int, float)):
                    ws.write(data_row, c, val, cell_style)
                else:
                    ws.write(data_row, c, str(val), cell_style)
            except Exception as _e:
                log.debug("_write_label_block data row col=%d err=%s", c, _e)
                ws.write(data_row, c, str(val), cell_style)


def labels_preview(
    routes: list[dict],
    file_type: str,
    products_ref: list | None,
    departments_ref: list | None,
    only_product: str | None = None,
    only_dept_key: str | None = None,
) -> list[tuple[str, str, int]]:
    """
    Возвращает список (продукт, отдел, кол-во маршрутов) для предпросмотра этикеток.
    Без записи файлов. only_product / only_dept_key — фильтр по одному продукту или отделу.
    """
    active = [r for r in routes if not r.get("excluded")]

    def include_product(prod_name: str) -> bool:
        if not products_ref or not departments_ref:
            return True
        prod = next((p for p in products_ref if p.get("name") == prod_name), None)
        if not prod:
            return True
        dept_key = prod.get("deptKey")
        if not dept_key:
            return False
        if only_dept_key and dept_key != only_dept_key:
            return False
        for dept in departments_ref:
            if dept.get("key") == dept_key:
                if not dept.get("labelsEnabled", True):
                    return False
                if dept.get("labelsFor", "both") in ("both", file_type):
                    return True
                return False
            for sub in dept.get("subdepts", []):
                if sub.get("key") == dept_key:
                    if not sub.get("labelsEnabled", True):
                        return False
                    if sub.get("labelsFor", "both") in ("both", file_type):
                        return True
                    return False
        return True

    products_with_templates: dict[str, str] = {}
    for p in products_ref or []:
        tpl = p.get("labelTemplatePath") or ""
        if not tpl or not os.path.isfile(tpl):
            continue
        name = p.get("name", "")
        if only_product and name != only_product:
            continue
        if include_product(name):
            products_with_templates[name] = p.get("deptKey") or ""

    result: list[tuple[str, str, int]] = []
    for prod_name, dept_key in products_with_templates.items():
        route_nums: set[str] = set()
        for route in active:
            for prod in route.get("products", []):
                if prod.get("name") == prod_name:
                    route_nums.add(str(route.get("routeNum", "")))
                    break
        dept_name = _dept_display_name(dept_key, departments_ref)
        result.append((prod_name, dept_name, len(route_nums)))
    return sorted(result, key=lambda x: (x[1], x[0]))


def generate_labels_from_templates(
    routes: list[dict],
    output_dir: str,
    file_type: str,
    products_ref: list | None,
    departments_ref: list | None,
    only_product: str | None = None,
    only_dept_key: str | None = None,
) -> list[str]:
    """
    Создаёт этикетки XLS по шаблонам продуктов.
    Один файл на продукт (или два для сыпучки: до 4 кг / после 4 кг). Количество сравнивается как float.
    Учитывается labelsEnabled отдела/подотдела. Режимы: default; chistchenka (макс. кг на этикетку, дубликаты пока не останется < макс — последняя этикетка; файл: продукт_дата_основной/увеличение);
    сыпучка — два файла до/после порога.
    only_product / only_dept_key — генерация только для одного продукта или отдела.
    """
    active = [r for r in routes if not r.get("excluded")]
    created: list[str] = []
    styles = _get_styles()

    def include_product(prod_name: str) -> bool:
        if not products_ref or not departments_ref:
            return True
        prod = next((p for p in products_ref if p.get("name") == prod_name), None)
        if not prod:
            return True
        dept_key = prod.get("deptKey")
        if not dept_key:
            return False  # этикетки только для продуктов с отделом
        if only_dept_key and dept_key != only_dept_key:
            return False
        for dept in departments_ref:
            if dept.get("key") == dept_key:
                if not dept.get("labelsEnabled", True):
                    return False
                if dept.get("labelsFor", "both") in ("both", file_type):
                    return True
                return False
            for sub in dept.get("subdepts", []):
                if sub.get("key") == dept_key:
                    if not sub.get("labelsEnabled", True):
                        return False
                    if sub.get("labelsFor", "both") in ("both", file_type):
                        return True
                    return False
        return True

    active_sorted = sorted(active, key=_label_sort_key_route)

    prod_map = {p["name"]: p for p in (products_ref or [])}
    _apply_pcs(active_sorted, prod_map)

    products_with_templates: dict[str, tuple[str, str, list]] = {}
    for p in products_ref or []:
        tpl = p.get("labelTemplatePath") or ""
        if not tpl or not os.path.isfile(tpl):
            continue
        name = p.get("name", "")
        if only_product and name != only_product:
            continue
        if include_product(name):
            layout = p.get("labelLayout")  # list of {row, col, field} or None
            if layout and not isinstance(layout, list):
                layout = None
            products_with_templates[name] = (tpl, p.get("deptKey") or "", layout or [])

    def _write_product_labels(
        item_list: list[tuple[str, str, float | None]],
        save_path: str,
        template_rows: int,
        ncols: int,
        matrix: list,
        label_layout: list[dict] | None = None,
    ) -> None:
        wb = xlwt.Workbook(encoding="utf-8")
        ws = wb.add_sheet("Этикетки")
        block_height = template_rows + 1  # строки шаблона + одна строка с данными
        page_breaks: list[int] = []
        row = 0
        for route_num, house, qty in item_list:
            _write_label_block(ws, matrix, template_rows, ncols, row, route_num, house, qty, styles, label_layout)
            row += block_height
            page_breaks.append(row)  # разрыв страницы перед следующим блоком
        if page_breaks:
            ws.horz_page_breaks = [(r, 0, 255) for r in page_breaks[:-1]]  # не после последнего блока
        wb.save(save_path)

    for prod_name, (template_path, dept_key, label_layout) in products_with_templates.items():
        mode = _label_print_mode_for_dept(dept_key, departments_ref)
        label_rules = _label_rules_for_dept(dept_key, departments_ref)
        try:
            nrows, ncols, matrix, last_filled = _load_template_matrix(template_path)
        except Exception as _e:
            log.warning("Не удалось загрузить шаблон '%s': %s", template_path, _e)
            continue

        template_rows = nrows if nrows > 0 else 1
        ncols = max(ncols, 3)

        items: list[tuple[str, str, float | None]] = []
        for route in active_sorted:
            route_num = str(route.get("routeNum", ""))
            house = extract_house_number(route.get("address", ""))
            for prod in route.get("products", []):
                if prod.get("name") != prod_name:
                    continue
                items.append((route_num, house, prod.get("displayQuantity", prod.get("quantity"))))

        if not items:
            continue

        safe_name = _safe_filename(prod_name)

        if mode == "sypuchka":
            syp = label_rules.get("sypuchka") or {}
            threshold = float(syp.get("thresholdKg", 4))
            label_below = syp.get("labelBelow", "меньше 4 кг")
            label_above = syp.get("labelAbove", "больше 4 кг")
            items_below = [(rn, h, q) for rn, h, q in items if q is not None and float(q) <= threshold]
            items_above = [(rn, h, q) for rn, h, q in items if q is None or float(q) > threshold]
            for title_suffix, item_list in [(label_below, items_below), (label_above, items_above)]:
                if not item_list:
                    continue
                fname = f"Этикетки {title_suffix} {safe_name}.xls"
                save_path = os.path.join(output_dir, fname)
                cnt = 0
                while os.path.exists(save_path):
                    cnt += 1
                    save_path = os.path.join(output_dir, f"Этикетки {title_suffix} {safe_name}_{cnt}.xls")
                _write_product_labels(item_list, save_path, template_rows, ncols, matrix, label_layout)
                created.append(save_path)

        elif mode == "chistchenka":
            ch = label_rules.get("chistchenka") or {}
            max_kg = float(ch.get("maxKgPerLabel", 5))
            if max_kg <= 0:
                max_kg = 5.0
            expanded: list[tuple[str, str, float]] = []
            for route_num, house, qty in items:
                try:
                    val = float(qty) if qty is not None else 0.0
                except (TypeError, ValueError):
                    val = 0.0
                while val > 0:
                    take = min(max_kg, val)
                    expanded.append((route_num, house, take))
                    val -= take
            date_str = _format_date(date.today())
            type_lbl_short = "основной" if file_type == "main" else "увеличение"
            fname = f"{safe_name}_{date_str}_{type_lbl_short}.xls"
            save_path = os.path.join(output_dir, fname)
            cnt = 0
            while os.path.exists(save_path):
                cnt += 1
                save_path = os.path.join(output_dir, f"{safe_name}_{date_str}_{type_lbl_short}_{cnt}.xls")
            _write_product_labels(expanded, save_path, template_rows, ncols, matrix, label_layout)
            created.append(save_path)

        else:
            date_str = _format_date(date.today())
            type_lbl_short = "основной" if file_type == "main" else "увеличение"
            fname = f"{safe_name}_{date_str}_{type_lbl_short}.xls"
            save_path = os.path.join(output_dir, fname)
            cnt = 0
            while os.path.exists(save_path):
                cnt += 1
                save_path = os.path.join(output_dir, f"{safe_name}_{date_str}_{type_lbl_short}_{cnt}.xls")
            _write_product_labels(items, save_path, template_rows, ncols, matrix, label_layout)
            created.append(save_path)

    return created


def _fmt_qty_with_pcs(prod: dict) -> str:
    """Форматирует количество с опциональным значением шт.
    Использует displayQuantity (уже с учётом множителя замены), если задано.
    """
    qty = prod.get("displayQuantity", prod.get("quantity"))
    unit = prod.get("unit", "")
    pcs = prod.get("pcs")
    if qty is None:
        return ""
    qty_str = f"{qty} {unit}".strip() if unit else str(qty)
    if pcs is not None:
        return f"{qty_str} / {pcs} шт"
    return qty_str


# ─────────────────────────── Общие маршруты ───────────────────────────────

def generate_general_routes(
    routes: list[dict],
    file_type: str,
    save_path: str,
    products_settings: dict[str, dict],
) -> str:
    """
    Создаёт файл «Общие маршруты».
    Один лист на каждый маршрут, имя листа = номер маршрута.

    Структура листа (новый формат):
      Строка 1: дата + тип файла (заголовок)
      Строка 2: «№ маршрута» | «Адрес» (объединённая строка-заголовок таблицы)
                 значение: route_num | address
      Строка 3+: строки продуктов:
                 пусто | Название продукта | Ед. изм. | Количество [| Шт]

    Маршруты отсортированы по убыванию номера.
    """
    _apply_pcs(routes, products_settings)
    routes_sorted = _sort_routes(routes)

    wb = xlwt.Workbook(encoding="utf-8")
    styles = _get_styles()
    used_names: set[str] = set()

    date_str = _format_date(_tomorrow())
    type_lbl = _type_label(file_type)
    header_text = f"{date_str}  {type_lbl}"

    for route in routes_sorted:
        sheet_name = _unique_sheet_name(str(route.get("routeNum", "?")), used_names)
        ws = wb.add_sheet(sheet_name)
        _apply_page_margins(ws, for_labels=False)

        products = route.get("products", [])
        has_pcs = any(p.get("pcs") is not None for p in products)
        n_prods = len(products)
        n_data_cols = 5 if has_pcs else 4  # пусто | продукт | ед.изм. | кол-во [| шт]

        # ── Строка 1: дата + тип файла ──────────────────────────────────────
        ws.write_merge(0, 0, 0, n_data_cols - 1, header_text, styles["title"])

        # ── Строка 2: заголовок таблицы — номер маршрута и адрес ────────────
        # Колонки: [№ маршрута] [Адрес] объединены на всю ширину
        route_num_str = str(route.get("routeNum", ""))
        address = route.get("address", "")

        ws.write(1, 0, route_num_str, styles["header"])
        # Адрес занимает оставшиеся столбцы (объединяем со 2-й по последнюю)
        if n_data_cols > 2:
            ws.write_merge(1, 1, 1, n_data_cols - 1, address, styles["header_wrap"])
        else:
            ws.write(1, 1, address, styles["header_wrap"])

        # ── Строка 3: заголовки столбцов продуктов ──────────────────────────
        prod_headers = ["Продукт", "Ед. изм.", "Количество"]
        if has_pcs:
            prod_headers.append("Шт")
        ws.write(2, 0, "", styles["header"])  # пустая ячейка под номер
        for ci, h in enumerate(prod_headers):
            ws.write(2, 1 + ci, h, styles["header"])

        # ── Строки 4+: данные продуктов ─────────────────────────────────────
        if n_prods == 0:
            # Нет продуктов — одна пустая строка
            ws.write(3, 0, "", styles["cell"])
            ws.write(3, 1, "", styles["cell"])
            ws.write(3, 2, "", styles["cell"])
            ws.write(3, 3, "", styles["num"])
            if has_pcs:
                ws.write(3, 4, "", styles["num"])
        else:
            for pi, prod in enumerate(products):
                row = 3 + pi
                ws.write(row, 0, "", styles["cell"])  # пустая ячейка (под номер маршрута)
                ws.write(row, 1, prod.get("name", ""), styles["cell"])
                ws.write(row, 2, prod.get("unit", ""), styles["cell"])
                qty = prod.get("displayQuantity", prod.get("quantity"))
                ws.write(row, 3, qty if qty is not None else "", styles["num"])
                if has_pcs:
                    pcs = prod.get("pcs")
                    ws.write(row, 4, pcs if pcs is not None else "", styles["num"])

        # ── Ширина столбцов ──────────────────────────────────────────────────
        _set_col_width(ws, 0, 14)   # № маршрута / пустая
        _set_col_width(ws, 1, 42)   # Адрес / Продукт
        _set_col_width(ws, 2, 12)   # Ед. изм.
        _set_col_width(ws, 3, 14)   # Количество
        if has_pcs:
            _set_col_width(ws, 4, 8)  # Шт

    wb.save(save_path)
    return save_path


# ─────────────────────────── Форматы файлов по отделам ────────────────────

def _write_dept_wide(
    ws: xlwt.Worksheet,
    routes: list[dict],
    dept_name: str,
    date_str: str,
    type_lbl: str,
    styles: dict[str, xlwt.XFStyle],
    sort_asc: bool = False,
) -> None:
    """
    Формат 1 (wide): каждый уникальный продукт — отдельный столбец.

    Структура:
      Строка 1: заголовок (объединённая)
      Строка 2: Маршрут | Адрес | ПродуктA | ПродуктB | ...
      Строка 3+: данные маршрутов (одна строка на маршрут)

    Значение в ячейке продукта: "5 кг / 3 шт" (шт — опционально).
    """
    # Собираем уникальные продукты в порядке первого появления
    unique_prods: list[str] = []
    seen: set[str] = set()
    for route in routes:
        for prod in route.get("products", []):
            name = prod.get("name", "")
            if name and name not in seen:
                seen.add(name)
                unique_prods.append(name)

    n_prod_cols = len(unique_prods)
    total_cols = 2 + n_prod_cols  # Маршрут + Адрес + продукты

    # Строка 1: заголовок
    title = f"Маршруты по {dept_name} {date_str} {type_lbl}"
    if total_cols > 1:
        ws.write_merge(0, 0, 0, total_cols - 1, title, styles["title"])
    else:
        ws.write(0, 0, title, styles["title"])

    # Строка 2: заголовки столбцов
    ws.write(1, 0, "Маршрут", styles["header"])
    ws.write(1, 1, "Адрес", styles["header"])
    for ci, pname in enumerate(unique_prods):
        ws.write(1, 2 + ci, pname, styles["header"])

    # Данные
    routes_sorted = _sort_routes(routes, sort_asc)
    for ri, route in enumerate(routes_sorted):
        row = 2 + ri
        route_num_str = str(route.get("routeNum", ""))
        address = route.get("address", "")

        ws.write(row, 0, route_num_str, styles["cell"])
        ws.write(row, 1, address, styles["cell_wrap"])

        # Строим словарь продуктов маршрута для быстрого поиска
        prod_by_name: dict[str, dict] = {
            p.get("name", ""): p for p in route.get("products", [])
        }

        for ci, pname in enumerate(unique_prods):
            prod = prod_by_name.get(pname)
            if prod is not None:
                cell_val = _fmt_qty_with_pcs(prod)
            else:
                cell_val = ""
            ws.write(row, 2 + ci, cell_val, styles["cell"])

    # Ширина столбцов
    _set_col_width(ws, 0, 14)
    _set_col_width(ws, 1, 42)
    for ci in range(n_prod_cols):
        _set_col_width(ws, 2 + ci, 20)


def _write_dept_rows(
    ws: xlwt.Worksheet,
    routes: list[dict],
    dept_name: str,
    date_str: str,
    type_lbl: str,
    styles: dict[str, xlwt.XFStyle],
    sort_asc: bool = False,
) -> None:
    """
    Формат 2 (rows): строчный формат — строка маршрута + строки продуктов.

    Структура:
      Строка 1: заголовок (объединённая)
      Строка 2: Маршрут | Адрес | Кол-во | Шт
               (2-я строка заголовка: — | — | ед.изм. | шт)
      Строка 3+: для каждого маршрута:
        - строка маршрута: номер | адрес | — | —
        - строки продуктов: — | название продукта | количество | шт_значение

    Маршруты отсортированы по убыванию номера.
    """
    # Строка 1: заголовок
    ws.write_merge(0, 0, 0, 3, f"Маршруты по {dept_name} {date_str} {type_lbl}",
                   styles["title"])

    # Строка 2: заголовки (первая строка)
    ws.write(1, 0, "Маршрут",    styles["header"])
    ws.write(1, 1, "Адрес",      styles["header"])
    ws.write(1, 2, "Кол-во",     styles["header"])
    ws.write(1, 3, "Шт",         styles["header"])

    # Строка 3: вторая строка заголовков (единицы измерения)
    ws.write(2, 0, "",           styles["header"])
    ws.write(2, 1, "",           styles["header"])
    ws.write(2, 2, "ед. изм.",   styles["header"])
    ws.write(2, 3, "шт",         styles["header"])

    # Данные начинаются с строки 4 (индекс 3)
    routes_sorted = _sort_routes(routes, sort_asc)
    current_row = 3

    for route in routes_sorted:
        products = route.get("products", [])
        route_num_str = str(route.get("routeNum", ""))
        address = route.get("address", "")

        # Строка маршрута
        ws.write(current_row, 0, route_num_str, styles["cell"])
        ws.write(current_row, 1, address,       styles["cell_wrap"])
        ws.write(current_row, 2, "",            styles["cell"])
        ws.write(current_row, 3, "",            styles["cell"])
        current_row += 1

        # Строки продуктов
        for prod in products:
            pname = prod.get("name", "")
            qty   = prod.get("displayQuantity", prod.get("quantity"))
            pcs   = prod.get("pcs")
            unit  = prod.get("unit", "")

            qty_str = f"{qty} {unit}".strip() if (qty is not None and unit) else (str(qty) if qty is not None else "")
            pcs_str = str(pcs) if pcs is not None else ""

            ws.write(current_row, 0, "",       styles["cell"])
            ws.write(current_row, 1, pname,    styles["cell"])
            ws.write(current_row, 2, qty_str,  styles["cell"])
            ws.write(current_row, 3, pcs_str,  styles["cell"])
            current_row += 1

    # Ширина столбцов
    _set_col_width(ws, 0, 14)
    _set_col_width(ws, 1, 42)
    _set_col_width(ws, 2, 16)
    _set_col_width(ws, 3, 10)


# ─────────────────────────── Устаревший формат (совместимость) ────────────

AVAILABLE_COLS: dict[str, str] = {
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

DEFAULT_COLS: list[dict] = [
    {"field": "routeNumber", "label": None, "merged": False},
    {"field": "address",     "label": None, "merged": False},
    {"field": "product",     "label": None, "merged": False},
    {"field": "unit",        "label": None, "merged": False},
    {"field": "quantity",    "label": None, "merged": False},
]

_COL_WIDTHS: dict[str, int] = {
    "routeNumber": 14, "address": 42, "product": 32,
    "unit": 12, "quantity": 14, "pcs": 8, "productQty": 32,
    "productsWide": 20, "nomenclature": 42,
}


def _get_col_label(col: dict) -> str:
    if col.get("label"):
        return col["label"]
    if col.get("merged") and col.get("productName"):
        return col["productName"]
    return AVAILABLE_COLS.get(col["field"], col["field"])


def _get_template(dept_key: str, templates: list[dict]) -> dict | None:
    """Возвращает шаблон для отдела или None."""
    for tmpl in templates:
        if tmpl.get("deptKey") == dept_key:
            return tmpl
    return templates[0] if templates else None


def _get_template_cols(dept_key: str, templates: list[dict]) -> list[dict]:
    """Returns list of column dicts for the given dept (legacy)."""
    tmpl = _get_template(dept_key, templates)
    if tmpl:
        cols = tmpl.get("columns", [])
        if cols:
            if isinstance(cols[0], str):
                return [{"field": c, "label": None, "merged": False} for c in cols]
            return cols
    return DEFAULT_COLS[:]


def _get_template_format(dept_key: str, templates: list[dict]) -> str:
    """Возвращает формат шаблона: 'wide', 'rows', или '' (legacy)."""
    tmpl = _get_template(dept_key, templates)
    if tmpl:
        return tmpl.get("format", "")
    return ""


def _resolve_merged_cols(
    template_cols: list[dict],
    routes: list[dict],
) -> list[dict]:
    """
    Автоопределение объединённых столбцов (productQty).
    """
    unique_products: list[str] = []
    seen: set[str] = set()
    for route in routes:
        for prod in route.get("products", []):
            name = prod.get("name", "")
            if name and name not in seen:
                seen.add(name)
                unique_products.append(name)

    result: list[dict] = []
    for col in template_cols:
        if col["field"] == "productsWide":
            for pname in unique_products:
                result.append({
                    "field": "productQty",
                    "label": pname,
                    "merged": True,
                    "productName": pname,
                })
        elif col["field"] == "productQty":
            if col.get("productName"):
                result.append(col)
            elif len(unique_products) == 1:
                result.append({
                    "field": "productQty",
                    "label": unique_products[0],
                    "merged": True,
                    "productName": unique_products[0],
                })
            else:
                for pname in unique_products:
                    result.append({
                        "field": "productQty",
                        "label": pname,
                        "merged": True,
                        "productName": pname,
                    })
        else:
            result.append(col)
    return result


def _write_dept_sheet_nomenclature(
    ws: xlwt.Worksheet,
    routes: list[dict],
    dept_name: str,
    date_str: str,
    type_lbl: str,
    template_cols: list[dict],
    styles: dict[str, xlwt.XFStyle],
    sort_asc: bool = False,
) -> None:
    """
    Запись по шаблону с колонкой «Номенклатура»: заголовок «Номенклатура»,
    в первой строке блока — адрес, в следующих — продукты отдела.
    № маршрута только в строке с адресом.
    """
    routes_sorted = _sort_routes(routes, sort_asc)
    n_cols = len(template_cols)
    title = f"Маршруты по {dept_name} {date_str} {type_lbl}"
    if n_cols > 1:
        ws.write_merge(0, 0, 0, n_cols - 1, title, styles["title"])
    else:
        ws.write(0, 0, title, styles["title"])

    for ci, col_def in enumerate(template_cols):
        lbl = "Номенклатура" if col_def.get("field") == "nomenclature" else _get_col_label(col_def)
        ws.write(1, ci, lbl, styles["header"])

    current_row = 2
    for route in routes_sorted:
        products = route.get("products", [])
        route_num_str = str(route.get("routeNum", ""))
        address = route.get("address", "")

        for pi in range(1 + len(products)):
            row = current_row + pi
            is_address_row = pi == 0
            prod = products[pi - 1] if pi > 0 else None

            for ci, col_def in enumerate(template_cols):
                field = col_def["field"]
                if field == "routeNumber":
                    ws.write(row, ci, route_num_str if is_address_row else "", styles["cell"])
                elif field == "address":
                    ws.write(row, ci, address if is_address_row else "", styles["cell_wrap"])
                elif field == "nomenclature":
                    if is_address_row:
                        ws.write(row, ci, address, styles["cell_wrap"])
                    elif prod is not None:
                        cell_val = _fmt_qty_with_pcs(prod)
                        ws.write(row, ci, f"{prod.get('name', '')} {cell_val}".strip(), styles["cell"])
                    else:
                        ws.write(row, ci, "", styles["cell"])
                elif field == "productQty" and col_def.get("merged") and col_def.get("productName"):
                    if prod is not None and prod.get("name") == col_def.get("productName"):
                        qty = prod.get("displayQuantity", prod.get("quantity"))
                        ws.write(row, ci, qty if qty is not None else "", styles["num"])
                    else:
                        ws.write(row, ci, "", styles["cell"])
                elif field == "product" and prod is not None:
                    ws.write(row, ci, prod.get("name", ""), styles["cell"])
                elif field == "unit" and prod is not None:
                    ws.write(row, ci, prod.get("unit", ""), styles["cell"])
                elif field == "quantity" and prod is not None:
                    qty = prod.get("displayQuantity", prod.get("quantity"))
                    ws.write(row, ci, qty if qty is not None else "", styles["num"])
                elif field == "pcs" and prod is not None:
                    ws.write(row, ci, prod.get("pcs") if prod.get("pcs") is not None else "", styles["num"])
                else:
                    ws.write(row, ci, "", styles["cell"])

        current_row += 1 + len(products)

    for ci, col_def in enumerate(template_cols):
        _set_col_width(ws, ci, _COL_WIDTHS.get(col_def["field"], 16))


def _write_dept_sheet(
    ws: xlwt.Worksheet,
    routes: list[dict],
    dept_name: str,
    date_str: str,
    type_lbl: str,
    template_cols: list[dict],
    styles: dict[str, xlwt.XFStyle],
    sort_asc: bool = False,
) -> None:
    """Legacy: записывает данные на лист отдела по column-based шаблону."""
    template_cols = _resolve_merged_cols(template_cols, routes)
    if any(c.get("field") == "nomenclature" for c in template_cols):
        _write_dept_sheet_nomenclature(
            ws, routes, dept_name, date_str, type_lbl, template_cols, styles, sort_asc
        )
        return
    n_cols = len(template_cols)

    title = f"Маршруты по {dept_name} {date_str} {type_lbl}"
    if n_cols > 1:
        ws.write_merge(0, 0, 0, n_cols - 1, title, styles["title"])
    else:
        ws.write(0, 0, title, styles["title"])

    for ci, col_def in enumerate(template_cols):
        ws.write(1, ci, _get_col_label(col_def), styles["header"])

    routes_sorted = _sort_routes(routes, sort_asc)
    current_row = 2

    for route in routes_sorted:
        products = route.get("products", [])
        n_prods = max(len(products), 1)
        route_num_str = str(route.get("routeNum", ""))
        address = route.get("address", "")

        for pi in range(n_prods):
            prod = products[pi] if pi < len(products) else {}
            row = current_row + pi

            for ci, col_def in enumerate(template_cols):
                field = col_def["field"]
                merged = col_def.get("merged", False)

                if field == "routeNumber":
                    if pi == 0:
                        if n_prods > 1:
                            ws.write_merge(row, row + n_prods - 1, ci, ci,
                                           route_num_str, styles["cell"])
                        else:
                            ws.write(row, ci, route_num_str, styles["cell"])

                elif field == "address":
                    if pi == 0:
                        if n_prods > 1:
                            ws.write_merge(row, row + n_prods - 1, ci, ci,
                                           address, styles["cell_wrap"])
                        else:
                            ws.write(row, ci, address, styles["cell_wrap"])

                elif field == "product":
                    ws.write(row, ci, prod.get("name", ""), styles["cell"])

                elif field == "unit":
                    ws.write(row, ci, prod.get("unit", ""), styles["cell"])

                elif field == "quantity":
                    qty = prod.get("displayQuantity", prod.get("quantity"))
                    ws.write(row, ci, qty if qty is not None else "", styles["num"])

                elif field == "pcs":
                    pcs = prod.get("pcs")
                    ws.write(row, ci, pcs if pcs is not None else "", styles["num"])

                elif field == "productQty" and merged:
                    target_name = col_def.get("productName", "")
                    target_prod = next(
                        (p for p in products if p.get("name", "") == target_name),
                        None
                    )
                    if target_prod is not None:
                        qty = target_prod.get("displayQuantity", target_prod.get("quantity"))
                        ws.write(row, ci, qty if qty is not None else "", styles["num"])
                    elif pi == 0:
                        ws.write(row, ci, "", styles["num"])

        current_row += n_prods

    for ci, col_def in enumerate(template_cols):
        _set_col_width(ws, ci, _COL_WIDTHS.get(col_def["field"], 16))


# ─────────────────────────── Публичные функции генерации ──────────────────

def _write_dept_by_format(
    ws: xlwt.Worksheet,
    routes: list[dict],
    dept_name: str,
    date_str: str,
    type_lbl: str,
    fmt: str,
    template_cols: list[dict],
    styles: dict[str, xlwt.XFStyle],
    sort_asc: bool = False,
) -> None:
    """Выбирает нужную функцию записи в зависимости от формата шаблона."""
    if fmt == "wide":
        _write_dept_wide(ws, routes, dept_name, date_str, type_lbl, styles, sort_asc)
    elif fmt == "rows":
        _write_dept_rows(ws, routes, dept_name, date_str, type_lbl, styles, sort_asc)
    else:
        # Legacy: column-based шаблон
        _write_dept_sheet(ws, routes, dept_name, date_str, type_lbl,
                          template_cols, styles, sort_asc)


def generate_single_dept_file(
    group: dict,
    file_type: str,
    save_path: str,
    prod_map: dict[str, dict],
    templates: list[dict],
    sort_asc: bool = False,
) -> str:
    """
    Создаёт файл для одного отдела/подотдела.

    Args:
        group: {"key", "name", "routes": [...]}
        file_type: "main" | "increase"
        save_path: полный путь к файлу .xls
        prod_map: dict {name: product_dict}
        templates: список шаблонов из data_store
        sort_asc: True — по возрастанию, False — по убыванию
    """
    routes = [r for r in group["routes"] if not r.get("excluded", False)]
    _apply_pcs(routes, prod_map)

    fmt = _get_template_format(group["key"], templates)
    template_cols = _get_template_cols(group["key"], templates)
    date_str = _format_date(_tomorrow())
    type_lbl = _type_label(file_type)

    wb = xlwt.Workbook(encoding="utf-8")
    styles = _get_styles()

    sheet_name = _safe_sheet_name(group["name"])
    ws = wb.add_sheet(sheet_name)
    _apply_page_margins(ws, for_labels=False)

    _write_dept_by_format(ws, routes, group["name"], date_str, type_lbl,
                          fmt, template_cols, styles, sort_asc)

    wb.save(save_path)
    return save_path


def generate_dept_files(
    dept_groups: list[dict],
    file_type: str,
    save_dir: str,
    prod_map: dict[str, dict],
    templates: list[dict],
    sort_asc: bool = False,
) -> list[str]:
    """
    Создаёт файлы для всех отделов/подотделов.

    Args:
        dept_groups: список {"key", "name", "routes": [...]}
        file_type: "main" | "increase"
        save_dir: папка сохранения
        prod_map: dict {name: product_dict}
        templates: список шаблонов
        sort_asc: True — по возрастанию, False — по убыванию

    Returns:
        Список путей к созданным файлам.
    """
    date_str = _format_date(_tomorrow())
    type_suffix = _type_suffix(file_type)
    type_lbl = _type_label(file_type)

    created: list[str] = []
    styles = _get_styles()

    for group in dept_groups:
        safe_name = _safe_filename(group["name"])
        filename = f"Маршруты {safe_name} {date_str} {type_suffix}.xls"
        save_path = os.path.join(save_dir, filename)

        if os.path.exists(save_path):
            base = os.path.splitext(save_path)[0]
            counter = 2
            while os.path.exists(save_path):
                save_path = f"{base}_{counter}.xls"
                counter += 1

        routes = [r for r in group["routes"] if not r.get("excluded", False)]
        _apply_pcs(routes, prod_map)

        fmt = _get_template_format(group["key"], templates)
        template_cols = _get_template_cols(group["key"], templates)

        wb = xlwt.Workbook(encoding="utf-8")
        ws = wb.add_sheet(_safe_sheet_name(group["name"]))
        _apply_page_margins(ws, for_labels=False)
        _write_dept_by_format(ws, routes, group["name"], date_str, type_lbl,
                              fmt, template_cols, styles, sort_asc)
        wb.save(save_path)
        created.append(save_path)

    return created
