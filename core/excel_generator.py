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

import math
import os
import re
from datetime import date, timedelta
from typing import Any

import xlwt

ROUTE_SIGN = "\u2116"

# ─────────────────────────── Кэш стилей ──────────────────────────────────

_STYLES: dict[str, xlwt.XFStyle] | None = None


def _get_styles() -> dict[str, xlwt.XFStyle]:
    """Возвращает кэшированный набор стилей (создаётся один раз)."""
    global _STYLES
    if _STYLES is not None:
        return _STYLES

    font_bold = xlwt.Font()
    font_bold.bold = True
    font_bold.height = 200  # 10pt

    font_normal = xlwt.Font()
    font_normal.height = 200

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
    return _STYLES


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
    Рассчитывает количество в штуках.

    Алгоритм:
      whole = floor(quantity / pcs_per_unit)
      remainder = quantity - whole * pcs_per_unit
      half = pcs_per_unit / 2
      round_up:   remainder >= half → +1
      round_down: remainder >  half → +1
    """
    if pcs_per_unit <= 0:
        return 0
    whole = math.floor(quantity / pcs_per_unit)
    remainder = quantity - whole * pcs_per_unit
    half = pcs_per_unit / 2
    return whole + (1 if (remainder >= half if round_up else remainder > half) else 0)


def _apply_pcs(routes: list[dict], prod_map: dict[str, dict]) -> list[dict]:
    """
    Добавляет поле pcs к продуктам маршрутов.
    prod_map: {name: product_settings_dict} — передаётся снаружи.
    Округление берётся по категории маршрута (ШК/СД): roundUpШК, roundUpСД; при отсутствии — roundUp.
    Изменяет продукты in-place (не копирует маршруты).
    """
    for route in routes:
        route_cat = route.get("routeCategory") or "ШК"
        for prod in route.get("products", []):
            sp = prod_map.get(prod["name"])
            pcs = None
            if sp and sp.get("showPcs") and prod.get("unit", "").lower() != "шт":
                qty = prod.get("quantity")
                if qty is not None:
                    try:
                        if route_cat == "СД":
                            round_up = sp.get("roundUpСД") if "roundUpСД" in sp else sp.get("roundUp", True)
                        else:
                            round_up = sp.get("roundUpШК") if "roundUpШК" in sp else sp.get("roundUp", True)
                        pcs = calc_pcs(
                            float(qty),
                            float(sp.get("pcsPerUnit", 1)),
                            bool(round_up),
                        )
                    except (ValueError, TypeError):
                        pass
            prod["pcs"] = pcs
    return routes


def _unique_sheet_name(name: str, used: set[str]) -> str:
    """Генерирует уникальное имя листа (макс 31 символ)."""
    base = str(name)[:28]
    candidate = base
    counter = 2
    while candidate in used:
        candidate = f"{base}_{counter}"
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
    Извлекает номер дома, строения или цифры перед символом № (U+2116) из адреса для этикетки.
    Порядок поиска: д.5, стр.2, строение 3, корп.1, цифры перед №.
    """
    if not address:
        return ""
    s = str(address).strip()
    # д.5, д. 5, д.5а, д.109/1
    m = re.search(r"д\.\s*(\d+(?:/\d+)?[а-яА-Яa-zA-Z]*)", s, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    # стр.2, стр. 2
    m = re.search(r"стр\.\s*(\d+)", s, re.IGNORECASE)
    if m:
        return m.group(1)
    # строение 3, строен. 3
    m = re.search(r"строен\.?\s*(\d+)", s, re.IGNORECASE)
    if m:
        return m.group(1)
    # корп.1
    m = re.search(r"корп\.\s*(\d+)", s, re.IGNORECASE)
    if m:
        return m.group(1)
    # цифры перед символом № (U+2116)
    m = re.search(r"(\d+)\s*[№\u2116]", s)
    if m:
        return m.group(1)
    return ""


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


def _route_sort_key_labels(route_num: str) -> tuple[int, str]:
    """Ключ сортировки маршрутов по возрастанию для этикеток."""
    try:
        return (1, int(str(route_num).strip()), str(route_num))
    except (ValueError, TypeError):
        return (0, 0, str(route_num))


def _load_template_matrix(template_path: str) -> tuple[int, int, list]:
    """Загружает шаблон XLS в матрицу (nrows, ncols, list of rows). Каждая ячейка — значение для записи."""
    import xlrd
    wb = xlrd.open_workbook(template_path, formatting_info=False)
    sheet = wb.sheet_by_index(0)
    nrows, ncols = sheet.nrows, sheet.ncols
    matrix = []
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
    return nrows, ncols, matrix


def _write_label_block(ws: Any, matrix: list, nrows: int, ncols: int, start_row: int,
                      route_num: str, house: str, qty_val: Any, styles: dict) -> None:
    """
    Пишет один блок этикетки: копия шаблона (matrix) с последней строкой = маршрут, дом, количество.
    Столбец 1 (номер маршрута) — жёлтый фон. Столбец 2 — дом/строение/цифры перед №. Столбец 3 — количество.
    """
    style_yellow = styles.get("cell_yellow", styles["cell"])
    for r in range(nrows):
        for c in range(ncols):
            val = matrix[r][c] if r < len(matrix) and c < len(matrix[r]) else ""
            if r == nrows - 1 and c >= 0:
                if c == 0:
                    val = route_num
                elif c == 1:
                    val = house
                elif c == 2:
                    val = qty_val if qty_val is not None else ""
            if r == nrows - 1 and c == 0:
                cell_style = style_yellow
            elif isinstance(val, (int, float)):
                cell_style = styles.get("num", styles["cell"])
            else:
                cell_style = styles["cell"]
            try:
                if isinstance(val, (int, float)):
                    ws.write(start_row + r, c, val, cell_style)
                else:
                    ws.write(start_row + r, c, str(val), cell_style)
            except Exception:
                ws.write(start_row + r, c, str(val), cell_style)


def generate_labels_from_templates(
    routes: list[dict],
    output_dir: str,
    file_type: str,
    products_ref: list | None,
    departments_ref: list | None,
) -> list[str]:
    """
    Создаёт этикетки XLS по шаблонам продуктов.
    Один файл на продукт (или два для сыпучки: до 4 кг / после 4 кг). Количество сравнивается как float.
    Учитывается labelsEnabled отдела/подотдела. Режимы: default; chistchenka (≤5 кг на этикетку);
    сыпучка — два файла до 4 кг / после 4 кг.
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
            return True
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

    active_sorted = sorted(active, key=lambda r: _route_sort_key_labels(str(r.get("routeNum", ""))))

    products_with_templates: dict[str, tuple[str, str]] = {}
    for p in products_ref or []:
        tpl = p.get("labelTemplatePath") or ""
        if tpl and os.path.isfile(tpl) and include_product(p.get("name", "")):
            products_with_templates[p.get("name", "")] = (tpl, p.get("deptKey") or "")

    for prod_name, (template_path, dept_key) in products_with_templates.items():
        mode = _label_print_mode_for_dept(dept_key, departments_ref)
        try:
            nrows, ncols, matrix = _load_template_matrix(template_path)
        except Exception:
            continue

        items: list[tuple[str, str, float | None]] = []
        for route in active_sorted:
            route_num = str(route.get("routeNum", ""))
            house = extract_house_number(route.get("address", ""))
            for prod in route.get("products", []):
                if prod.get("name") != prod_name:
                    continue
                items.append((route_num, house, prod.get("quantity")))

        if not items:
            continue

        safe_name = _safe_filename(prod_name)

        if mode == "sypuchka":
            items_before = [(rn, h, q) for rn, h, q in items if q is not None and float(q) < 4]
            items_after = [(rn, h, q) for rn, h, q in items if q is None or float(q) >= 4]
            for title_suffix, item_list in [("до 4 кг", items_before), ("после 4 кг", items_after)]:
                if not item_list:
                    continue
                fname = f"Этикетки Сыпучка {title_suffix} {safe_name}.xls"
                save_path = os.path.join(output_dir, fname)
                cnt = 0
                while os.path.exists(save_path):
                    cnt += 1
                    save_path = os.path.join(output_dir, f"Этикетки Сыпучка {title_suffix} {safe_name}_{cnt}.xls")
                wb = xlwt.Workbook(encoding="utf-8")
                ws = wb.add_sheet("Этикетки")
                row = 0
                for route_num, house, qty in item_list:
                    _write_label_block(ws, matrix, nrows, ncols, row, route_num, house, qty, styles)
                    row += nrows + 1
                wb.save(save_path)
                created.append(save_path)

        elif mode == "chistchenka":
            MAX_KG = 5.0
            expanded: list[tuple[str, str, float]] = []
            for route_num, house, qty in items:
                try:
                    val = float(qty) if qty is not None else 0.0
                except (TypeError, ValueError):
                    val = 0.0
                if val <= MAX_KG:
                    expanded.append((route_num, house, val))
                else:
                    rest = val
                    while rest > 0:
                        part = min(MAX_KG, rest)
                        expanded.append((route_num, house, part))
                        rest -= part
            fname = f"Этикетки_{safe_name}.xls"
            save_path = os.path.join(output_dir, fname)
            cnt = 0
            while os.path.exists(save_path):
                cnt += 1
                save_path = os.path.join(output_dir, f"Этикетки_{safe_name}_{cnt}.xls")
            wb = xlwt.Workbook(encoding="utf-8")
            ws = wb.add_sheet("Этикетки")
            row = 0
            for route_num, house, qty in expanded:
                _write_label_block(ws, matrix, nrows, ncols, row, route_num, house, qty, styles)
                row += nrows + 1
            wb.save(save_path)
            created.append(save_path)

        else:
            fname = f"Этикетки_{safe_name}.xls"
            save_path = os.path.join(output_dir, fname)
            cnt = 0
            while os.path.exists(save_path):
                cnt += 1
                save_path = os.path.join(output_dir, f"Этикетки_{safe_name}_{cnt}.xls")
            wb = xlwt.Workbook(encoding="utf-8")
            ws = wb.add_sheet("Этикетки")
            row = 0
            for route_num, house, qty in items:
                _write_label_block(ws, matrix, nrows, ncols, row, route_num, house, qty, styles)
                row += nrows + 1
            wb.save(save_path)
            created.append(save_path)

    return created


def _fmt_qty_with_pcs(prod: dict) -> str:
    """Форматирует количество с опциональным значением шт.
    Результат: "5 кг / 3 шт" или "5 кг" если шт не задано.
    """
    qty = prod.get("quantity")
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
                qty = prod.get("quantity")
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
            qty   = prod.get("quantity")
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
    "routeNumber": "№ маршрута",
    "address":     "Адрес",
    "product":     "Продукт",
    "unit":        "Ед. изм.",
    "quantity":    "Количество",
    "pcs":         "Шт",
    "productQty":  "Продукт (кол-во)",
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
        if col["field"] == "productQty":
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
                    qty = prod.get("quantity")
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
                        qty = target_prod.get("quantity")
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

    sheet_name = group["name"][:31]
    ws = wb.add_sheet(sheet_name)

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
        ws = wb.add_sheet(group["name"][:31])
        _write_dept_by_format(ws, routes, group["name"], date_str, type_lbl,
                              fmt, template_cols, styles, sort_asc)
        wb.save(save_path)
        created.append(save_path)

    return created
