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

import copy
import json
import logging
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import date, timedelta, datetime
import time
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


def _dept_special_mode_raw(name: str, explicit_mode: str | None = None) -> str:
    mode = (explicit_mode or "").strip().lower()
    if mode:
        if mode in ("chistchenka", "sypuchka", "polufabricates", "polufabrikaty"):
            return "polufabricates" if mode in ("polufabricates", "polufabrikaty") else mode
        return mode
    lname = (name or "").strip().lower()
    if "чищен" in lname:
        return "chistchenka"
    if "полуфаб" in lname:
        return "polufabricates"
    if "сыпуч" in lname:
        return "sypuchka"
    return "default"


def _build_dept_mode_map() -> dict[str, str]:
    mode_map: dict[str, str] = {}
    for dept in data_store.get_ref("departments") or []:
        dkey = str(dept.get("key") or "")
        if dkey:
            mode_map[dkey] = _dept_special_mode_raw(
                str(dept.get("name") or ""),
                str(dept.get("labelPrintMode") or ""),
            )
        for sub in dept.get("subdepts", []):
            skey = str(sub.get("key") or "")
            if skey:
                mode_map[skey] = _dept_special_mode_raw(
                    str(sub.get("name") or ""),
                    str(sub.get("labelPrintMode") or ""),
                )
    return mode_map


def get_dept_special_mode(dept_key: str | None) -> str:
    """Возвращает спец-режим отдела: default/chistchenka/sypuchka/polufabricates."""
    if not dept_key:
        return "default"
    return _build_dept_mode_map().get(str(dept_key), "default")


def get_routes_date_str() -> str:
    """Дата маршрутов (как в файлах): обычно следующий день."""
    return _format_date(_tomorrow())


def get_routes_day_folder(base_dir: str, date_str: str | None = None) -> str:
    date_part = date_str or get_routes_date_str()
    return os.path.join(base_dir, f"Маршруты {date_part}")


def get_routes_type_folder(base_dir: str, file_type: str, date_str: str | None = None) -> str:
    day_dir = get_routes_day_folder(base_dir, date_str)
    sub = "Увеличение" if file_type == "increase" else "Основные"
    return os.path.join(day_dir, sub)


def get_general_routes_path(base_dir: str, file_type: str, date_str: str | None = None) -> str:
    date_part = date_str or get_routes_date_str()
    type_dir = get_routes_type_folder(base_dir, file_type, date_part)
    return os.path.join(type_dir, f"Общие маршруты {date_part}.xls")


def get_dept_routes_path(base_dir: str, file_type: str, dept_name: str, date_str: str | None = None) -> str:
    type_dir = get_routes_type_folder(base_dir, file_type, date_str)
    safe_name = _safe_filename(dept_name)
    folder = os.path.join(type_dir, f"Маршруты {safe_name}")
    return os.path.join(folder, f"Маршруты {safe_name}.xls")


def calc_pcs(quantity: float, pcs_per_unit: float, round_up: bool = True) -> int:
    """
    Рассчитывает количество штук (старая логика: ceil/floor по всему отношению).
    """
    if pcs_per_unit <= 0 or quantity <= 0:
        return 0
    ratio = quantity / pcs_per_unit
    if round_up:
        return max(0, int(math.ceil(ratio)))
    return max(0, int(math.floor(ratio)))


def calc_pcs_tail(
    quantity: float,
    pcs_per_unit: float,
    round_tail_from: float,
) -> int:
    """
    Рассчитывает штуки по логике «хвостик ШК/СД».

    Хвостик — граница остатка (в тех же единицах, что количество: кг, л и т.д.).
    При остатке от этой границы включительно округление в большую сторону (+1 шт).
    Пример: 1 шт = 0.7 кг, хвостик ШК = 0.35 → при остатке ≥ 0.35 кг добавляем 1 шт.

    Целые штуки = floor(quantity / pcs_per_unit).
    Остаток = quantity - целые * pcs_per_unit.
    round_tail_from = 0: при любом ненулевом остатке +1 шт (как ceil).
    round_tail_from > 0: при остатке >= round_tail_from +1 шт.
    """
    if pcs_per_unit <= 0 or quantity <= 0:
        return 0
    full = int(math.floor(quantity / pcs_per_unit))
    remainder = quantity - full * pcs_per_unit
    if round_tail_from <= 0:
        add_one = remainder > 0
    else:
        add_one = remainder >= round_tail_from
    if add_one:
        return full + 1
    return full


def _apply_pcs(routes: list[dict], prod_map: dict[str, dict]) -> list[dict]:
    """
    Добавляет к продуктам маршрутов displayQuantity (с учётом множителя замены) и pcs.
    prod_map: {name: product_settings_dict}. Коэффициент замены (quantityMultiplier), напр. 1.25
    для пересчёта очищенных → грязные: отображаемое количество = количество × коэффициент.
    Округление берётся по категории маршрута (ШК/СД), а для некоторых учреждений —
    в большую сторону по проценту отдела (см. is_always_round_up_institution, get_institution_round_percent).
    """
    dept_mode_map = _build_dept_mode_map()
    for route in routes:
        route_cat = route.get("routeCategory") or "ШК"
        addr = route.get("address", "")
        force_round_up = is_always_round_up_institution(addr)
        for prod in route.get("products", []):
            sp = prod_map.get(prod["name"], {})
            dept_key = sp.get("deptKey")
            dept_mode = dept_mode_map.get(str(dept_key or ""), "default")
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
            pcs_tail = None
            if sp.get("showPcs") and prod.get("unit", "").lower() != "шт":
                eff_qty = prod.get("displayQuantity")
                if eff_qty is not None:
                    try:
                        val = float(eff_qty)
                        pcu = float(sp.get("pcsPerUnit", 1))
                        if pcu <= 0:
                            pcs = 0
                        else:
                            if dept_mode == "polufabricates":
                                # Для полуфабрикатов округления нет: показываем целые шт + хвостик.
                                pcs = max(0, int(math.floor(val / pcu)))
                                pcs_tail = max(0.0, float(val - pcs * pcu))
                            else:
                                # Порог: ниже — 0 шт, от порога и выше — расчёт по округлению
                                min_qty = sp.get("minQtyForPcs")
                                if min_qty is not None and min_qty > 0:
                                    threshold = float(min_qty)
                                else:
                                    unit_lower = (prod.get("unit") or "").strip().lower()
                                    threshold = 0.2 if unit_lower in ("кг", "л", "kg", "l") else 0
                                if val < threshold:
                                    pcs = 0
                                else:
                                    if force_round_up:
                                        # Для заданных учреждений — в большую по % отдела
                                        pct = get_institution_round_percent(dept_key)
                                        round_tail = pcu * (pct / 100.0)
                                        pcs = calc_pcs_tail(val, pcu, round_tail)
                                    else:
                                        round_tail = (
                                            sp.get("roundTailFromСД") if route_cat == "СД" else sp.get("roundTailFromШК")
                                        )
                                        if round_tail is not None:
                                            round_tail = float(round_tail)
                                            pcs = calc_pcs_tail(val, pcu, round_tail)
                                        else:
                                            round_up = (
                                                sp.get("roundUpСД") if "roundUpСД" in sp else sp.get("roundUp", True)
                                                if route_cat == "СД"
                                                else sp.get("roundUpШК") if "roundUpШК" in sp else sp.get("roundUp", True)
                                            )
                                            pcs = calc_pcs(val, pcu, bool(round_up))
                    except (ValueError, TypeError):
                        pass
            prod["pcs"] = pcs
            prod["pcsTail"] = pcs_tail
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


def _sort_routes(routes: list[dict], sort_asc: bool = True) -> list[dict]:
    """Сортирует маршруты по номеру маршрута (числовая сортировка).
    sort_asc=True (по умолчанию) — по возрастанию.
    sort_asc=False — по убыванию.
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


def is_always_round_up_institution(address: str) -> bool:
    """
    Проверяет, включён ли адрес в список учреждений с округлением шт в большую сторону.
    Ключ учреждения: первые 3–4 цифры (109/1 → 109, 1391/2 → 1391).
    Адреса из excludeRoundUpAddresses исключаются.
    """
    key = data_store.get_institution_key_from_address(address or "")
    if not key:
        return False
    try:
        codes = data_store.get_setting("alwaysRoundUpInstitutions") or []
        if key not in codes:
            return False
        excluded = set(data_store.get_setting("excludeRoundUpAddresses") or [])
        if (address or "").strip() in excluded:
            return False
        return True
    except Exception:
        return False


def get_institution_round_percent(dept_key: str | None) -> float:
    """
    % от 1 шт для округления в учреждениях. Зависит от отдела (roundUpPercentByDept).
    """
    return data_store.get_round_up_percent_for_dept(dept_key)


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


def _load_template_matrix(
    template_path: str,
) -> tuple[int, int, list, int, list[int], list[int], list[tuple[int, int, int, int]], list[int]]:
    """
    Загружает шаблон XLS в матрицу.
    Скрытые строки (hidden=True или height=0) пропускаются полностью.
    Пустые строки в конце обрезаются.

    Возвращает:
        nrows       — число строк в обрезанной матрице
        ncols       — число столбцов
        matrix      — list[list] значений ячеек
        last_filled — индекс последней непустой строки
        row_heights — высоты строк (1/20 пункта) для каждой видимой строки матрицы
        col_widths  — ширины столбцов (1/256 символа)
        merges      — объединённые ячейки в пространстве матрицы:
                      list of (r1, r2, c1, c2) включительно
        source_rows — исходные номера строк шаблона (0-based) для каждой строки матрицы
    """
    import xlrd
    # formatting_info=True нужен для rowinfo_map, colinfo_map, merged_cells
    wb = xlrd.open_workbook(template_path, formatting_info=True)
    sheet = wb.sheet_by_index(0)
    ncols = sheet.ncols
    matrix: list[list] = []
    row_heights: list[int] = []
    source_rows: list[int] = []
    last_filled = -1

    # Карта: исходный индекс строки файла → индекс в матрице (только видимые строки)
    orig_to_matrix: dict[int, int] = {}

    for r in range(sheet.nrows):
        ri = sheet.rowinfo_map.get(r)
        if ri is not None and (ri.hidden or ri.height == 0):
            continue
        orig_to_matrix[r] = len(matrix)
        row: list = []
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
        row_heights.append(ri.height if ri is not None else 0)
        source_rows.append(r)
        if any(v != "" for v in row):
            last_filled = len(matrix) - 1

    # Обрезаем хвостовые пустые строки
    nrows_trimmed = (last_filled + 1) if last_filled >= 0 else (1 if matrix else 0)
    matrix = matrix[:nrows_trimmed]
    row_heights = row_heights[:nrows_trimmed]
    source_rows = source_rows[:nrows_trimmed]

    # Ширины столбцов
    col_widths: list[int] = []
    for c in range(ncols):
        ci = sheet.colinfo_map.get(c)
        col_widths.append(ci.width if ci is not None else 0)

    # Объединённые ячейки — перевести в пространство матрицы
    # xlrd: merged_cells = list of (row_lo, row_hi_excl, col_lo, col_hi_excl)
    merges: list[tuple[int, int, int, int]] = []
    for (row_lo, row_hi, col_lo, col_hi) in (sheet.merged_cells or []):
        visible = [orig_to_matrix[r] for r in range(row_lo, row_hi) if r in orig_to_matrix]
        if not visible:
            continue
        m_r1, m_r2 = visible[0], visible[-1]
        # Отсекаем всё, что выходит за пределы обрезанной матрицы
        if m_r1 >= nrows_trimmed:
            continue
        m_r2 = min(m_r2, nrows_trimmed - 1)
        # Переводим col_hi из «exclusive» в «inclusive»
        merges.append((m_r1, m_r2, col_lo, col_hi - 1))

    return len(matrix), ncols, matrix, last_filled, row_heights, col_widths, merges, source_rows


def load_label_template_matrix(template_path: str) -> tuple[int, int, list, int]:
    """
    Публичная обёртка для загрузки шаблона этикетки (предпросмотр в UI).
    Возвращает (nrows, ncols, matrix, last_filled_row).
    """
    nrows, ncols, matrix, last_filled, _rh, _cw, _mg, _sr = _load_template_matrix(template_path)
    return nrows, ncols, matrix, last_filled


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
    row_heights: list[int] | None = None,
    merges: list[tuple[int, int, int, int]] | None = None,
) -> None:
    """
    Пишет один блок этикетки: копия первых template_rows строк шаблона (matrix),
    затем строка с данными (route_num / house / qty).

    row_heights — высоты строк шаблона (1/20 пункта).
    merges      — объединённые ячейки в пространстве матрицы: (r1, r2, c1, c2) включительно.
                  Ячейки внутри объединения (не верхний-левый угол) пропускаются при одиночной
                  записи; само объединение пишется через ws.write_merge().
    """
    style_yellow = styles.get("cell_yellow", styles["cell"])
    style_num = styles.get("num", styles["cell"])

    # --- Строим множество ячеек, покрытых объединением (кроме верхнего-левого угла) ---
    merged_covered: set[tuple[int, int]] = set()
    if merges:
        for (r1, r2, c1, c2) in merges:
            for mr in range(r1, r2 + 1):
                for mc in range(c1, c2 + 1):
                    if mr != r1 or mc != c1:
                        merged_covered.add((mr, mc))

    # --- Копируем строки шаблона (0 .. template_rows-1) ---
    for r in range(template_rows):
        # Применяем высоту строки из шаблона
        if row_heights and r < len(row_heights) and row_heights[r] > 0:
            ws_row = ws.row(start_row + r)
            ws_row.height = row_heights[r]
            ws_row.height_mismatch = True
        for c in range(ncols):
            if (r, c) in merged_covered:
                continue  # будет записано write_merge ниже
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

    # --- Записываем объединённые ячейки шаблона ---
    if merges:
        for (r1, r2, c1, c2) in merges:
            if r1 >= template_rows:
                continue
            r2_clamped = min(r2, template_rows - 1)
            val = matrix[r1][c1] if r1 < len(matrix) and c1 < len(matrix[r1]) else ""
            cell_style = style_num if isinstance(val, (int, float)) else styles["cell"]
            r1_abs = start_row + r1
            r2_abs = start_row + r2_clamped
            try:
                if r1_abs == r2_abs and c1 == c2:
                    # Одиночная ячейка — обычная запись
                    ws.write(r1_abs, c1, val if isinstance(val, (int, float)) else str(val), cell_style)
                elif isinstance(val, (int, float)):
                    ws.write_merge(r1_abs, r2_abs, c1, c2, val, cell_style)
                else:
                    ws.write_merge(r1_abs, r2_abs, c1, c2, str(val), cell_style)
            except Exception as _e:
                log.debug("_write_label_block merge r=%d-%d c=%d-%d err=%s", r1, r2, c1, c2, _e)

    # --- Заполнение данных: по layout или по умолчанию (строка данных, колонки 0,1,2) ---
    data_row = start_row + template_rows
    ncols_write = max(ncols, 3)
    if label_layout:
        values_by_cell: dict[tuple[int, int], Any] = {}
        for pl in label_layout:
            r, c = pl.get("row", template_rows), pl.get("col", 0)
            f = pl.get("field")
            if f == "routeNumber":
                val, cell_style = route_num, style_yellow
            elif f == "house":
                val, cell_style = house, styles["cell"]
            elif f == "quantity":
                val, cell_style = (qty_val if qty_val is not None else ""), style_num
            else:
                continue

            key = (r, c)
            if key in values_by_cell:
                prev_val, prev_style = values_by_cell[key]
                if prev_val and val:
                    new_val = f"{prev_val} {val}"
                else:
                    new_val = val or prev_val
                new_style = style_yellow if (prev_style is style_yellow or cell_style is style_yellow) else prev_style
                values_by_cell[key] = (new_val, new_style)
            else:
                values_by_cell[key] = (val, cell_style)

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


def _build_label_cell_values(
    route_num: str,
    house: str,
    qty_val: Any,
    label_layout: list[dict] | None,
    template_rows: int,
) -> tuple[dict[tuple[int, int], Any], int]:
    """
    Возвращает значения для подстановки в этикетку и число добавочных строк.
    Ключ словаря: (row_idx, col_idx) в пространстве шаблона.
    """
    style_values: dict[tuple[int, int], Any] = {}
    placements = label_layout or []
    if placements:
        for pl in placements:
            r = int(pl.get("row", template_rows))
            c = int(pl.get("col", 0))
            field = pl.get("field")
            if field == "routeNumber":
                val = route_num
            elif field == "house":
                val = house
            elif field == "quantity":
                val = qty_val if qty_val is not None else ""
            else:
                continue
            key = (r, c)
            prev = style_values.get(key)
            if prev not in (None, "") and val not in (None, ""):
                style_values[key] = f"{prev} {val}"
            else:
                style_values[key] = val if val not in (None, "") else (prev or "")
    else:
        style_values = {
            (template_rows, 0): route_num,
            (template_rows, 1): house,
            (template_rows, 2): qty_val if qty_val is not None else "",
        }
    max_row_idx = max((r for r, _c in style_values), default=template_rows - 1)
    extra_rows = max(0, max_row_idx - template_rows + 1)
    return style_values, extra_rows


def _strip_windows_zone_identifier(path: str) -> None:
    """
    Удаляет ADS-метку Zone.Identifier у файла на Windows.
    Это снижает шанс открытия в Protected View и зависания при "Разрешить редактирование".
    """
    if os.name != "nt":
        return
    try:
        ads_path = f"{os.path.abspath(path)}:Zone.Identifier"
        if os.path.exists(os.path.abspath(path)):
            os.remove(ads_path)
    except Exception:
        # best-effort, генерацию не валим
        pass


def _try_generate_labels_exact_excel(
    template_path: str,
    item_list: list[tuple[str, str, float | None]],
    save_path: str,
    template_rows: int,
    source_rows: list[int],
    label_layout: list[dict] | None = None,
) -> bool:
    """
    Точное создание этикеток через установленный Excel в отдельном подпроцессе.
    Это защищает основное приложение от зависаний COM/скрытых диалогов Excel.
    """
    if not item_list:
        return True
    worker_path = os.path.join(os.path.dirname(__file__), "excel_exact_worker.py")
    if not os.path.isfile(worker_path):
        log.warning("Не найден worker точной генерации: %s", worker_path)
        return False

    temp_dir = tempfile.mkdtemp(prefix="labels_excel_job_")
    payload_path = os.path.join(temp_dir, "payload.json")
    try:
        with open(payload_path, "w", encoding="utf-8") as fh:
            json.dump({
                "template_path": template_path,
                "item_list": item_list,
                "save_path": save_path,
                "template_rows": template_rows,
                "source_rows": source_rows,
                "label_layout": label_layout or [],
            }, fh, ensure_ascii=False)

        timeout_sec = max(120, min(1200, 45 + len(item_list) * 4))
        for attempt in (1, 2):
            try:
                proc = subprocess.run(
                    [sys.executable, worker_path, payload_path],
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    timeout=timeout_sec,
                )
                if proc.returncode == 0 and os.path.isfile(save_path):
                    return True
                if proc.stdout.strip():
                    log.warning("Excel worker stdout (attempt %s): %s", attempt, proc.stdout.strip())
                if proc.stderr.strip():
                    log.warning("Excel worker stderr (attempt %s): %s", attempt, proc.stderr.strip())
            except subprocess.TimeoutExpired:
                log.warning(
                    "Excel worker timeout (attempt %s, timeout=%ss, items=%s)",
                    attempt, timeout_sec, len(item_list)
                )
            if attempt == 1:
                # Вторая попытка в отдельном чистом процессе Excel.
                continue
        return False
    except Exception as exc:
        log.warning("Не удалось запустить Excel worker: %s", exc)
        return False
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def _try_export_xls_to_pdf(xls_path: str, pdf_path: str) -> bool:
    """Экспортирует XLS в PDF через установленный Excel в отдельном подпроцессе."""
    worker_path = os.path.join(os.path.dirname(__file__), "excel_pdf_worker.py")
    if not os.path.isfile(worker_path):
        return False
    try:
        proc = subprocess.run(
            [sys.executable, worker_path, xls_path, pdf_path],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=60,
        )
        if proc.returncode == 0 and os.path.isfile(pdf_path):
            return True
        if proc.stdout.strip():
            log.warning("PDF worker stdout: %s", proc.stdout.strip())
        if proc.stderr.strip():
            log.warning("PDF worker stderr: %s", proc.stderr.strip())
        return False
    except Exception as exc:
        log.warning("Не удалось экспортировать XLS в PDF: %s", exc)
        return False


def _finalize_label_output(
    xls_path: str,
    output_format: str,
    created: list[str],
) -> None:
    """
    Добавляет итоговые файлы в список created в зависимости от output_format:
    - xls: только xls
    - pdf: пытается сделать pdf, при неудаче оставляет xls
    - both: xls + pdf (если удалось)
    """
    fmt = (output_format or "xls").lower()
    if fmt not in ("xls", "pdf", "both"):
        fmt = "xls"

    _strip_windows_zone_identifier(xls_path)
    if fmt == "xls":
        created.append(xls_path)
        return

    pdf_path = os.path.splitext(xls_path)[0] + ".pdf"
    if _try_export_xls_to_pdf(xls_path, pdf_path):
        _strip_windows_zone_identifier(pdf_path)
        if fmt == "pdf":
            try:
                os.remove(xls_path)
            except Exception:
                pass
            created.append(pdf_path)
            return
        created.append(xls_path)
        created.append(pdf_path)
        return

    # Если PDF не получилось — не теряем результат.
    created.append(xls_path)


def _append_labels_diagnostics(
    output_dir: str,
    lines: list[str],
) -> None:
    """Пишет диагностический отчёт генерации этикеток."""
    try:
        os.makedirs(output_dir, exist_ok=True)
        report_path = os.path.join(output_dir, "_labels_generation_report.txt")
        with open(report_path, "a", encoding="utf-8") as fh:
            fh.write("\n")
            fh.write("=" * 90 + "\n")
            fh.write(f"Запуск: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n")
            for line in lines:
                fh.write(line.rstrip() + "\n")
    except Exception:
        pass


def _run_excel_exact_worker(payload: dict, timeout_sec: int = 180) -> tuple[int, str, str]:
    """Запускает excel_exact_worker.py с payload и возвращает (code, stdout, stderr)."""
    worker_path = os.path.join(os.path.dirname(__file__), "excel_exact_worker.py")
    if not os.path.isfile(worker_path):
        return (2, "", f"worker not found: {worker_path}")
    temp_dir = tempfile.mkdtemp(prefix="excel_exact_job_")
    payload_path = os.path.join(temp_dir, "payload.json")
    try:
        with open(payload_path, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, ensure_ascii=False)
        proc = subprocess.run(
            [sys.executable, worker_path, payload_path],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout_sec,
        )
        return (proc.returncode, proc.stdout or "", proc.stderr or "")
    except Exception as exc:
        return (3, "", str(exc))
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def get_excel_printers() -> list[str]:
    """Возвращает список доступных принтеров для печати через Excel."""
    code, out, _err = _run_excel_exact_worker({"mode": "printers"}, timeout_sec=30)
    if code != 0:
        return []
    try:
        payload = json.loads(out.strip() or "{}")
        printers = payload.get("printers") or []
        if isinstance(printers, list):
            return [str(p) for p in printers if str(p).strip()]
    except Exception:
        pass
    return []


def open_label_live_preview(xls_path: str) -> None:
    """Открывает live-preview в отдельном экземпляре Excel и ждёт закрытия книги."""
    code, _out, err = _run_excel_exact_worker(
        {"mode": "preview", "xls_path": os.path.abspath(xls_path)},
        timeout_sec=24 * 3600,
    )
    if code != 0:
        raise RuntimeError(err or "Не удалось открыть live-preview Excel.")


def print_label_file(
    xls_path: str,
    printer_name: str | None = None,
    margins: dict[str, float] | None = None,
) -> str:
    """
    Печатает файл этикеток через Excel COM.
    Возвращает фактически использованный принтер (если удалось определить).
    """
    payload = {
        "mode": "print",
        "xls_path": os.path.abspath(xls_path),
        "printer_name": printer_name or "",
        "margins": margins or {
            "top_cm": 2.0,
            "right_cm": 2.0,
            "bottom_cm": 0.0,
            "left_cm": 0.0,
        },
    }
    code, out, err = _run_excel_exact_worker(payload, timeout_sec=300)
    if code != 0:
        raise RuntimeError(err or "Не удалось отправить этикетки на печать.")
    try:
        data = json.loads((out or "").strip() or "{}")
        return str(data.get("used_printer") or printer_name or "")
    except Exception:
        return str(printer_name or "")


def prepare_label_temp_file(
    routes: list[dict],
    file_type: str,
    products_ref: list | None,
    departments_ref: list | None,
    product_name: str,
    dept_key: str | None = None,
) -> tuple[str, str]:
    """
    Готовит временный XLS для preview/print по одному продукту.
    Возвращает (xls_path, temp_dir). temp_dir удаляется вызывающей стороной.
    """
    temp_dir = tempfile.mkdtemp(prefix="labels_preview_")
    created = generate_labels_from_templates(
        routes=routes,
        output_dir=temp_dir,
        file_type=file_type,
        products_ref=products_ref,
        departments_ref=departments_ref,
        only_product=product_name,
        only_dept_key=dept_key,
        dept_subfolders=False,
        overwrite=True,
        output_format="xls",
        strict_exact=True,
        diagnostics_dir=None,
    )
    xls_files = [p for p in created if p.lower().endswith(".xls")]
    if not xls_files:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise RuntimeError("Не удалось подготовить временный файл этикеток для предпросмотра.")
    return (xls_files[0], temp_dir)


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


def labels_preview_rows(
    routes: list[dict],
    file_type: str,
    products_ref: list | None,
    departments_ref: list | None,
    only_product: str | None = None,
    only_dept_key: str | None = None,
) -> list[tuple[str, str, str, str, str]]:
    """
    Возвращает список строк этикеток для таблицы предпросмотра.
    Каждая строка: (№ маршрута, Адрес, Продукт, Отдел, Кол-во).
    Отсортировано по номеру маршрута по возрастанию.
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
        return False

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

    prod_map = {p["name"]: p for p in (products_ref or [])}
    _apply_pcs(active, prod_map)

    active_sorted = sorted(
        active,
        key=lambda r: _route_sort_key_labels(str(r.get("routeNum", ""))),
    )

    rows: list[tuple[str, str, str, str, str]] = []
    for route in active_sorted:
        route_num = str(route.get("routeNum", ""))
        address = (route.get("address") or "").strip()
        for prod in route.get("products", []):
            prod_name = prod.get("name", "")
            if prod_name not in products_with_templates:
                continue
            dept_key = products_with_templates[prod_name]
            dept_name = _dept_display_name(dept_key, departments_ref)
            qty = prod.get("displayQuantity", prod.get("quantity"))
            qty_str = str(qty) if qty is not None else ""
            rows.append((route_num, address, prod_name, dept_name, qty_str))

    return rows


def generate_labels_from_templates(
    routes: list[dict],
    output_dir: str,
    file_type: str,
    products_ref: list | None,
    departments_ref: list | None,
    only_product: str | None = None,
    only_dept_key: str | None = None,
    dept_subfolders: bool = False,
    overwrite: bool = True,
    output_format: str = "xls",
    strict_exact: bool = True,
    diagnostics_dir: str | None = None,
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
    diagnostics: list[str] = []

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

    active_sorted = sorted(
        active,
        key=lambda r: _route_sort_key_labels(str(r.get("routeNum", "")))
    )

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

    def _apply_col_widths(ws_sheet: Any, col_widths: list[int], ncols_cnt: int) -> None:
        """Устанавливает ширины столбцов из шаблона на листе."""
        for ci, w in enumerate(col_widths[:ncols_cnt]):
            if w > 0:
                ws_sheet.col(ci).width = w

    def _write_product_labels(
        item_list: list[tuple[str, str, float | None]],
        save_path: str,
        template_rows: int,
        ncols: int,
        matrix: list,
        label_layout: list[dict] | None = None,
        row_heights: list[int] | None = None,
        col_widths: list[int] | None = None,
        merges: list[tuple[int, int, int, int]] | None = None,
    ) -> None:
        wb = xlwt.Workbook(encoding="utf-8")
        block_height = template_rows + 1  # строки шаблона + одна строка с данными
        _XLS_MAX_ROWS = 65535

        sheet_num = 1
        ws = wb.add_sheet("Этикетки")
        if col_widths:
            _apply_col_widths(ws, col_widths, ncols)
        page_breaks: list[int] = []
        row = 0

        for route_num, house, qty in item_list:
            # Если следующий блок не помещается на текущий лист — создаём новый
            if row + block_height > _XLS_MAX_ROWS:
                if page_breaks:
                    ws.horz_page_breaks = [(r, 0, 255) for r in page_breaks[:-1]]
                sheet_num += 1
                ws = wb.add_sheet(f"Этикетки {sheet_num}")
                if col_widths:
                    _apply_col_widths(ws, col_widths, ncols)
                page_breaks = []
                row = 0

            _write_label_block(
                ws, matrix, template_rows, ncols, row,
                route_num, house, qty, styles, label_layout, row_heights, merges,
            )
            row += block_height
            page_breaks.append(row)

        if page_breaks:
            ws.horz_page_breaks = [(r, 0, 255) for r in page_breaks[:-1]]
        wb.save(save_path)

    try:
        for prod_name, (template_path, dept_key, label_layout) in products_with_templates.items():
            started = time.perf_counter()
            mode = _label_print_mode_for_dept(dept_key, departments_ref)
            label_rules = _label_rules_for_dept(dept_key, departments_ref)
            try:
                nrows, ncols, matrix, last_filled, row_heights, col_widths, merges, source_rows = _load_template_matrix(template_path)
            except Exception as _e:
                log.warning("Не удалось загрузить шаблон '%s': %s", template_path, _e)
                diagnostics.append(f"[ERROR] {prod_name}: шаблон не прочитан ({_e})")
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
                diagnostics.append(f"[SKIP] {prod_name}: нет данных для генерации")
                continue

            safe_name = _safe_filename(prod_name)

            if dept_subfolders and dept_key:
                dept_display = _dept_display_name(dept_key, departments_ref)
                safe_dept = _safe_filename(dept_display) if dept_display else "Отдел"
                actual_dir = os.path.join(output_dir, safe_dept)
            else:
                actual_dir = output_dir
            os.makedirs(actual_dir, exist_ok=True)

            def _unique_path(base_path: str, base_name: str, ext: str = ".xls") -> str:
                if overwrite:
                    return base_path
                cnt = 0
                p = base_path
                while os.path.exists(p):
                    cnt += 1
                    p = os.path.join(actual_dir, f"{base_name}_{cnt}{ext}")
                return p

            try:
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
                        base_name = f"Этикетки {title_suffix} {safe_name}"
                        save_path = _unique_path(os.path.join(actual_dir, f"{base_name}.xls"), base_name)
                        exact_ok = _try_generate_labels_exact_excel(
                            template_path, item_list, save_path, template_rows, source_rows, label_layout
                        )
                        if not exact_ok and strict_exact:
                            raise RuntimeError(
                                f"Excel COM не смог создать этикетки для '{prod_name}' ({title_suffix})."
                            )
                        if not exact_ok:
                            _write_product_labels(
                                item_list, save_path, template_rows, ncols, matrix,
                                label_layout, row_heights, col_widths, merges,
                            )
                        _finalize_label_output(save_path, output_format, created)

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
                    base_name = f"{safe_name}_{date_str}_{type_lbl_short}"
                    save_path = _unique_path(os.path.join(actual_dir, f"{base_name}.xls"), base_name)
                    exact_ok = _try_generate_labels_exact_excel(
                        template_path, expanded, save_path, template_rows, source_rows, label_layout
                    )
                    if not exact_ok and strict_exact:
                        raise RuntimeError(
                            f"Excel COM не смог создать этикетки для '{prod_name}'."
                        )
                    if not exact_ok:
                        _write_product_labels(
                            expanded, save_path, template_rows, ncols, matrix,
                            label_layout, row_heights, col_widths, merges,
                        )
                    _finalize_label_output(save_path, output_format, created)

                else:
                    date_str = _format_date(date.today())
                    type_lbl_short = "основной" if file_type == "main" else "увеличение"
                    base_name = f"{safe_name}_{date_str}_{type_lbl_short}"
                    save_path = _unique_path(os.path.join(actual_dir, f"{base_name}.xls"), base_name)
                    exact_ok = _try_generate_labels_exact_excel(
                        template_path, items, save_path, template_rows, source_rows, label_layout
                    )
                    if not exact_ok and strict_exact:
                        raise RuntimeError(
                            f"Excel COM не смог создать этикетки для '{prod_name}'."
                        )
                    if not exact_ok:
                        _write_product_labels(
                            items, save_path, template_rows, ncols, matrix,
                            label_layout, row_heights, col_widths, merges,
                        )
                    _finalize_label_output(save_path, output_format, created)

                elapsed = int((time.perf_counter() - started) * 1000)
                diagnostics.append(
                    f"[OK] {prod_name}: mode={mode}, items={len(items)}, time={elapsed}ms"
                )
            except Exception as exc:
                elapsed = int((time.perf_counter() - started) * 1000)
                diagnostics.append(
                    f"[ERROR] {prod_name}: mode={mode}, time={elapsed}ms, error={exc}"
                )
                log.warning("Ошибка генерации этикеток для '%s': %s", prod_name, exc)
                continue
    finally:
        _append_labels_diagnostics(output_dir, diagnostics)
        if diagnostics_dir and os.path.abspath(diagnostics_dir) != os.path.abspath(output_dir):
            _append_labels_diagnostics(diagnostics_dir, diagnostics)

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
        tail = prod.get("pcsTail")
        if tail is not None and unit:
            try:
                tval = float(tail)
                if tval > 1e-9:
                    tail_txt = str(int(tval)) if abs(tval - round(tval)) < 1e-9 else f"{tval:.3f}".rstrip("0").rstrip(".")
                    return f"{qty_str} / {pcs} шт + {tail_txt} {unit}"
            except (TypeError, ValueError):
                pass
        return f"{qty_str} / {pcs} шт"
    return qty_str


# ─────────────────────────── Общие маршруты ───────────────────────────────

def generate_general_routes(
    routes: list[dict],
    file_type: str,
    save_path: str,
    products_settings: dict[str, dict],
    sort_asc: bool = True,
) -> str:
    """
    Создаёт файл «Общие маршруты».
    Один лист, все маршруты подряд с разрывом страницы перед каждым следующим блоком.

    Структура блока на маршрут:
      Строка 1: дата + тип файла (заголовок)
      Строка 2: «№ маршрута» | «Адрес»
      Строка 3: заголовки столбцов продуктов
      Строки 4+: данные продуктов

    Маршруты отсортированы по возрастанию номера (или по sort_asc).
    """
    _apply_pcs(routes, products_settings)
    routes_sorted = _sort_routes(routes, sort_asc)

    # Для одного листа — общее число столбцов по всем маршрутам (если хоть у одного есть шт)
    has_pcs_any = any(
        any(p.get("pcs") is not None for p in r.get("products", []))
        for r in routes_sorted
    )
    n_data_cols = 5 if has_pcs_any else 4

    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("Маршруты")
    _apply_page_margins(ws, for_labels=False)
    styles = _get_styles()

    date_str = _format_date(_tomorrow())
    type_lbl = _type_label(file_type)
    header_text = f"{date_str}  {type_lbl}"

    prod_headers = ["Продукт", "Ед. изм.", "Количество"]
    if has_pcs_any:
        prod_headers.append("Шт")

    page_breaks: list[int] = []
    row = 0

    for route in routes_sorted:
        if row > 0:
            page_breaks.append(row)  # разрыв страницы перед блоком маршрута

        products = route.get("products", [])
        has_pcs = has_pcs_any  # используем общие столбцы
        n_prods = len(products)

        # Строка 1: дата + тип файла
        ws.write_merge(row, row, 0, n_data_cols - 1, header_text, styles["title"])
        row += 1

        route_num_str = str(route.get("routeNum", ""))
        address = route.get("address", "")
        ws.write(row, 0, route_num_str, styles["header"])
        if n_data_cols > 2:
            ws.write_merge(row, row, 1, n_data_cols - 1, address, styles["header_wrap"])
        else:
            ws.write(row, 1, address, styles["header_wrap"])
        row += 1

        ws.write(row, 0, "", styles["header"])
        for ci, h in enumerate(prod_headers):
            ws.write(row, 1 + ci, h, styles["header"])
        row += 1

        if n_prods == 0:
            ws.write(row, 0, "", styles["cell"])
            ws.write(row, 1, "", styles["cell"])
            ws.write(row, 2, "", styles["cell"])
            ws.write(row, 3, "", styles["num"])
            if has_pcs_any:
                ws.write(row, 4, "", styles["num"])
            row += 1
        else:
            for prod in products:
                ws.write(row, 0, "", styles["cell"])
                ws.write(row, 1, prod.get("name", ""), styles["cell"])
                ws.write(row, 2, prod.get("unit", ""), styles["cell"])
                qty = prod.get("displayQuantity", prod.get("quantity"))
                ws.write(row, 3, qty if qty is not None else "", styles["num"])
                if has_pcs_any:
                    pcs = prod.get("pcs")
                    ws.write(row, 4, pcs if pcs is not None else "", styles["num"])
                row += 1

    if page_breaks:
        ws.horz_page_breaks = [(r, 0, 255) for r in page_breaks]

    _set_col_width(ws, 0, 14)
    _set_col_width(ws, 1, 42)
    _set_col_width(ws, 2, 12)
    _set_col_width(ws, 3, 14)
    if has_pcs_any:
        _set_col_width(ws, 4, 8)

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
    date_str: str | None = None,
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
    date_str = date_str or get_routes_date_str()
    type_lbl = _type_label(file_type)

    created: list[str] = []
    styles = _get_styles()

    for group in dept_groups:
        save_path = get_dept_routes_path(save_dir, file_type, group["name"], date_str)
        os.makedirs(os.path.dirname(save_path), exist_ok=True)

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


def _aggregate_pcs_totals_by_product(
    routes: list[dict],
    prod_map: dict[str, dict],
) -> dict[tuple[str, str, str], float]:
    data = copy.deepcopy(routes or [])
    _apply_pcs(data, prod_map)
    totals: dict[tuple[str, str, str], float] = {}
    for route in data:
        if route.get("excluded"):
            continue
        for prod in route.get("products", []):
            name = prod.get("name") or ""
            if not name:
                continue
            settings = prod_map.get(name) or {}
            if not settings.get("showPcs"):
                continue
            dept_key = str(settings.get("deptKey") or "")
            unit = str(prod.get("unit") or settings.get("unit") or "")
            pcs = prod.get("pcs")
            if pcs is None:
                continue
            try:
                key = (dept_key, name, unit)
                totals[key] = totals.get(key, 0.0) + float(pcs)
            except (TypeError, ValueError):
                continue
    return totals


def generate_pcs_compare_report(
    report_path: str,
    main_routes: list[dict],
    increase_routes: list[dict],
    products_ref: list[dict],
) -> str:
    """
    Создаёт отчет по продуктам с showPcs:
    Отдел | Продукт | Ед.изм. | Шт (Основные) | Шт (Увеличение)
    """
    prod_map = {p.get("name"): dict(p) for p in (products_ref or []) if p.get("name")}
    main_totals = _aggregate_pcs_totals_by_product(main_routes, prod_map)
    inc_totals = _aggregate_pcs_totals_by_product(increase_routes, prod_map)

    enabled_products: list[tuple[str, str, str]] = []
    for p in (products_ref or []):
        name = p.get("name")
        if not name or not p.get("showPcs"):
            continue
        enabled_products.append((str(p.get("deptKey") or ""), name, str(p.get("unit") or "")))

    all_keys = sorted(
        set(enabled_products) | set(main_totals.keys()) | set(inc_totals.keys()),
        key=lambda x: (
            (data_store.get_department_display_name(x[0]) or "").lower(),
            x[1].lower(),
        ),
    )

    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("Отчет Шт")
    _apply_page_margins(ws, for_labels=False)
    styles = _get_styles()

    date_str = get_routes_date_str()
    ws.write_merge(0, 0, 0, 4, f"Отчет по Шт {date_str}", styles["title"])
    ws.write(1, 0, "Отдел / Подотдел", styles["header"])
    ws.write(1, 1, "Продукт", styles["header"])
    ws.write(1, 2, "Ед. изм.", styles["header"])
    ws.write(1, 3, "Шт (Основные)", styles["header"])
    ws.write(1, 4, "Шт (Увеличение)", styles["header"])

    row = 2
    for dept_key, name, unit in all_keys:
        dept_name = data_store.get_department_display_name(dept_key) if dept_key else "Без отдела"
        m = main_totals.get((dept_key, name, unit), 0.0)
        i = inc_totals.get((dept_key, name, unit), 0.0)
        ws.write(row, 0, dept_name, styles["cell"])
        ws.write(row, 1, name, styles["cell"])
        ws.write(row, 2, unit, styles["cell"])
        ws.write(row, 3, int(m) if abs(m - round(m)) < 1e-9 else round(m, 1), styles["num"])
        ws.write(row, 4, int(i) if abs(i - round(i)) < 1e-9 else round(i, 1), styles["num"])
        row += 1

    _set_col_width(ws, 0, 28)
    _set_col_width(ws, 1, 36)
    _set_col_width(ws, 2, 12)
    _set_col_width(ws, 3, 16)
    _set_col_width(ws, 4, 16)
    os.makedirs(os.path.dirname(report_path), exist_ok=True)
    wb.save(report_path)
    return report_path
