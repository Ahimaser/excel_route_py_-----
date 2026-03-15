"""
xls_parser.py — Парсер XLS файлов маршрутов.

Оптимизации:
- Предварительное кэширование всех значений листа в list-of-lists (один проход)
- Использование frozenset для route_row_set (быстрее set при многократных in-проверках)
- Компилированный regex (один раз при импорте)
- Избегание повторных вызовов sheet.cell() — работаем с кэшем
- Минимальное количество промежуточных объектов

Реальная структура файла:
- Строки 1–14: шапка (ТРЕБОВАНИЕ-НАКЛАДНАЯ, заголовки таблицы)
- Данные с 15-й строки (0-indexed: row 14)
- Строка маршрута: объединение col1 (clo=1, chi>2) — адрес маршрута
- Строка продукта: объединение col2 (clo=2, chi<=5) — название продукта
- col6 = единица измерения
- col8 = количество (9-й столбец Excel, индекс 8)
- Номер маршрута извлекается из адреса после символа № (U+2116)

Порядок обработки parse_file():
1. Открыть книгу с formatting_info=True
2. Разобрать merged_cells → route_rows, product_row_set
3. Построить кэш ячеек (один проход)
4. Извлечь маршруты и продукты из кэша
"""
from __future__ import annotations

import logging
import re
import itertools
import xlrd
from typing import Any

log = logging.getLogger("xls_parser")

ROUTE_SIGN = "\u2116"  # №

# Количество строк шапки (ТРЕБОВАНИЕ-НАКЛАДНАЯ, заголовки таблицы)
SKIP_HEADER_ROWS: int = 14

# Компилируем regex один раз при импорте
_RE_ROUTE_NUM = re.compile(r"(\d+)")
_XL_EMPTY = xlrd.XL_CELL_EMPTY
_XL_BLANK = xlrd.XL_CELL_BLANK
_XL_NUMBER = xlrd.XL_CELL_NUMBER


def extract_route_number(address: str) -> str:
    """
    Извлекает номер маршрута из адреса после символа №.
    Возвращает строку с числом или 'Номер маршрута не определен'.
    """
    idx = address.find(ROUTE_SIGN)
    if idx == -1:
        return "Номер маршрута не определен"
    m = _RE_ROUTE_NUM.search(address, idx + 1)
    return m.group(1) if m else "Номер маршрута не определен"


def _build_cell_cache(sheet: xlrd.sheet.Sheet) -> list[list[tuple[int, Any]]]:
    """
    Строит кэш всех ячеек листа в виде list[row][col] = (ctype, value).
    Один проход по листу вместо многократных sheet.cell() вызовов.
    """
    nrows = sheet.nrows
    ncols = sheet.ncols
    cache: list[list[tuple[int, Any]]] = []
    for r in range(nrows):
        row_data: list[tuple[int, Any]] = []
        for c in range(ncols):
            cell = sheet.cell(r, c)
            row_data.append((cell.ctype, cell.value))
        cache.append(row_data)
    return cache


def _cell_str_cached(cache: list[list[tuple[int, Any]]], row: int, col: int) -> str:
    """Возвращает строковое значение ячейки из кэша."""
    try:
        ctype, value = cache[row][col]
    except IndexError:
        return ""
    if ctype in (_XL_EMPTY, _XL_BLANK):
        return ""
    if ctype == _XL_NUMBER:
        iv = int(value)
        return str(iv) if value == iv else str(value)
    return str(value).strip()


def _find_footer_start_row(cache: list[list[tuple[int, Any]]], nrows: int) -> int | None:
    """
    Находит первую строку подвала, сканируя с конца файла.
    Маркеры: «итого», «итого:», «всего учетных». Проверяет столбцы A, B, C.
    """
    for r in range(nrows - 1, -1, -1):
        for col in (0, 1, 2):
            cell_val = _cell_str_cached(cache, r, col).strip()
            if not cell_val:
                continue
            lower = cell_val.lower()
            if "итого" in lower or "всего учетных" in lower:
                return r
    return None


def parse_file(file_path: str) -> dict[str, Any]:
    """
    Парсит XLS файл и возвращает список маршрутов.

    Возвращает:
    {
        "routes": [
            {
                "routeNum": "21",
                "address": "109/1 ДС  ул.Лобановский лес д.2  М №21",
                "products": [
                    {"name": "Апельсин", "unit": "кг", "quantity": 1.517}
                ]
            }
        ],
        "uniqueProducts": [{"name": "Апельсин", "unit": "кг"}, ...]
    }
    """
    # Шаг 1: открываем книгу с formatting_info=True — обязательно для merged_cells
    wb = xlrd.open_workbook(file_path, formatting_info=True)
    sheet = wb.sheet_by_index(0)
    nrows = sheet.nrows

    # Шаг 2: разбираем merged_cells ПЕРВЫМ делом
    # route_rows  — строки с адресом маршрута (объединение начинается с col1)
    # product_row_set — строки с названием продукта (объединение начинается с col2)
    route_rows: list[int] = []
    product_row_set: set[int] = set()

    for rlo, rhi, clo, chi in sheet.merged_cells:
        if clo == 1 and chi > 2:
            route_rows.append(rlo)
        elif clo == 2 and chi <= 5:
            product_row_set.add(rlo)

    route_rows.sort()

    # Пропуск шапки: первые SKIP_HEADER_ROWS строк (заголовки таблицы)
    route_rows = [r for r in route_rows if r >= SKIP_HEADER_ROWS]
    product_row_set = {r for r in product_row_set if r >= SKIP_HEADER_ROWS}

    # Автопропуск: учитываем только диапазон данных (от первой строки маршрута до последней строки с данными)
    if route_rows or product_row_set:
        data_start = min(route_rows) if route_rows else min(product_row_set)
        data_end = max(itertools.chain(route_rows, product_row_set))
        route_rows = [r for r in route_rows if data_start <= r <= data_end]
        product_row_set = {r for r in product_row_set if data_start <= r <= data_end}

    # Шаг 3: строим кэш ячеек (один проход по листу)
    cell_cache = _build_cell_cache(sheet)

    # Пропуск подвала: строки с «Итого:», «Всего учетных единиц:»
    footer_start = _find_footer_start_row(cell_cache, nrows)
    if footer_start is not None:
        route_rows = [r for r in route_rows if r < footer_start]
        product_row_set = {r for r in product_row_set if r < footer_start}

    routes: list[dict[str, Any]] = []
    unique_products: dict[str, str] = {}  # name -> unit

    for i, route_row in enumerate(route_rows):
        address = _cell_str_cached(cell_cache, route_row, 1).strip()
        if not address or "итого" in address.lower():
            continue
        route_num = extract_route_number(address) if address else "Номер маршрута не определен"

        next_route_row = route_rows[i + 1] if i + 1 < len(route_rows) else nrows

        products: list[dict[str, Any]] = []
        for prod_row in range(route_row + 1, next_route_row):
            if prod_row not in product_row_set:
                continue
            name = _cell_str_cached(cell_cache, prod_row, 2).strip()
            if not name:
                continue
            if "итого" in name.lower():
                continue
            unit    = _cell_str_cached(cell_cache, prod_row, 6)  # col6 = ед.изм.
            qty_str = _cell_str_cached(cell_cache, prod_row, 8)  # col8 = количество (9-й столбец Excel)

            quantity: float | None = None
            if qty_str:
                try:
                    quantity = float(qty_str.replace(",", "."))
                except ValueError:
                    pass

            products.append({"name": name, "unit": unit, "quantity": quantity})

            if name not in unique_products:
                unique_products[name] = unit

        routes.append({
            "routeNum": route_num,
            "address": address,
            "products": products,
        })

    # Освобождаем кэш явно (помогает GC при больших файлах)
    del cell_cache

    unique_list = [{"name": n, "unit": u} for n, u in unique_products.items()]
    return {"routes": routes, "uniqueProducts": unique_list}


def parse_files(
    file_paths: list[str],
    file_categories: list[str] | None = None,
) -> dict[str, Any]:
    """
    Парсит несколько XLS файлов и объединяет результаты.
    file_categories: если задан, список той же длины — категория «ШК» или «СД» для каждого файла;
      каждой маршруту из файла i присваивается routeCategory = file_categories[i].
    Уникальные продукты дедуплицируются по имени.
    Автоматически заменяет варианты написания на канонические через алиасы.
    """
    if not file_paths:
        return {"routes": [], "uniqueProducts": [], "errors": []}

    from core import data_store  # ленивый импорт для избежания циклического импорта
    aliases = data_store.get_aliases()  # {variant: canonical}

    all_routes: list[dict[str, Any]] = []
    all_unique: dict[str, str] = {}  # canonical_name -> unit
    errors: list[str] = []

    for i, fp in enumerate(file_paths):
        cat = (file_categories[i] if file_categories and i < len(file_categories) else None) or "ШК"
        try:
            result = parse_file(fp)
            for route in result["routes"]:
                for prod in route["products"]:
                    prod["name"] = aliases.get(prod["name"], prod["name"])
                route["routeCategory"] = cat
            all_routes.extend(result["routes"])
            for p in result["uniqueProducts"]:
                canonical = aliases.get(p["name"], p["name"])
                if canonical not in all_unique:
                    all_unique[canonical] = p["unit"]
        except (OSError, ValueError, KeyError, TypeError, xlrd.XLRDError) as e:
            log.warning("Ошибка парсинга %s: %s", fp, e)
            errors.append(f"{fp}: {e}")

    return {
        "routes": all_routes,
        "uniqueProducts": [{"name": n, "unit": u} for n, u in all_unique.items()],
        "errors": errors,
    }
