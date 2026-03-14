from __future__ import annotations

import json
import os
import shutil
import sys
import tempfile
import time


def _build_label_cell_values(
    route_num: str,
    house: str,
    qty_val,
    label_layout: list[dict] | None,
    template_rows: int,
) -> tuple[dict[tuple[int, int], object], int]:
    values: dict[tuple[int, int], object] = {}
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
            prev = values.get(key)
            if prev not in (None, "") and val not in (None, ""):
                values[key] = f"{prev} {val}"
            else:
                values[key] = val if val not in (None, "") else (prev or "")
    else:
        values = {
            (template_rows, 0): route_num,
            (template_rows, 1): house,
            (template_rows, 2): qty_val if qty_val is not None else "",
        }
    max_row_idx = max((r for r, _c in values), default=template_rows - 1)
    extra_rows = max(0, max_row_idx - template_rows + 1)
    return values, extra_rows


def _run_generate(payload: dict, excel) -> int:
    template_path = payload["template_path"]
    item_list = [tuple(item) for item in payload["item_list"]]
    save_path = payload["save_path"]
    template_rows = int(payload["template_rows"])
    source_rows = [int(v) for v in payload["source_rows"]]
    label_layout = payload.get("label_layout") or []

    src_wb = None
    dst_wb = None
    temp_dir = None
    try:
        temp_dir = tempfile.mkdtemp(prefix="labels_excel_")
        temp_template = os.path.join(temp_dir, "template.xls")
        temp_output = os.path.join(temp_dir, "output.xls")
        shutil.copyfile(template_path, temp_template)

        src_wb = excel.Workbooks.Open(
            os.path.abspath(temp_template),
            UpdateLinks=False,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
            Notify=False,
            AddToMru=False,
        )
        src_sheet = src_wb.Worksheets(1)

        src_sheet.Copy()
        dst_wb = excel.ActiveWorkbook
        template_sheet = dst_wb.Worksheets(1)
        template_sheet.Name = "_template"

        keep_rows = set(r + 1 for r in source_rows)
        used_start = template_sheet.UsedRange.Row
        used_end = used_start + template_sheet.UsedRange.Rows.Count - 1
        for row_no in range(used_end, used_start - 1, -1):
            if row_no not in keep_rows:
                template_sheet.Rows(row_no).Delete()

        template_rows = len(source_rows)
        first_route_num, first_house, first_qty = item_list[0]
        _first_values, extra_rows = _build_label_cell_values(
            first_route_num, first_house, first_qty, label_layout, template_rows
        )
        block_height = template_rows + extra_rows
        if extra_rows > 0:
            for idx in range(extra_rows):
                insert_at = template_rows + idx + 1
                template_sheet.Rows(insert_at).Insert()
                if template_rows > 0:
                    template_sheet.Rows(template_rows).Copy(template_sheet.Rows(insert_at))
                template_sheet.Rows(insert_at).ClearContents()

        block_last_col = template_sheet.UsedRange.Column + template_sheet.UsedRange.Columns.Count - 1
        block_range = template_sheet.Range(
            template_sheet.Cells(1, 1),
            template_sheet.Cells(block_height, block_last_col),
        )

        blocks_per_sheet = max(1, 65536 // max(1, block_height))
        chunks = [item_list[i:i + blocks_per_sheet] for i in range(0, len(item_list), blocks_per_sheet)]

        def _write_block_values(ws, start_row_1based: int, values: dict[tuple[int, int], object]) -> None:
            for (row_idx, col_idx), val in values.items():
                cell = ws.Cells(start_row_1based + row_idx, col_idx + 1)
                if bool(cell.MergeCells):
                    cell = cell.MergeArea.Cells(1, 1)
                cell.Value = val

        def _copy_template_block(ws, start_row_1based: int) -> None:
            block_range.Copy(ws.Cells(start_row_1based, 1))
            for row_idx in range(1, block_height + 1):
                try:
                    ws.Rows(start_row_1based + row_idx - 1).RowHeight = template_sheet.Rows(row_idx).RowHeight
                except Exception:
                    pass

        created_sheets = []
        for chunk_idx, chunk_items in enumerate(chunks, start=1):
            template_sheet.Copy(After=dst_wb.Worksheets(dst_wb.Worksheets.Count))
            ws = dst_wb.Worksheets(dst_wb.Worksheets.Count)
            ws.Name = "Этикетки" if chunk_idx == 1 else f"Этикетки {chunk_idx}"
            created_sheets.append(ws)

            if len(chunk_items) > 1:
                for block_idx in range(1, len(chunk_items)):
                    dst_row = block_idx * block_height + 1
                    _copy_template_block(ws, dst_row)

            for block_idx, (route_num, house, qty_val) in enumerate(chunk_items):
                values, _unused = _build_label_cell_values(
                    route_num, house, qty_val, label_layout, template_rows
                )
                start_row = block_idx * block_height + 1
                _write_block_values(ws, start_row, values)
                if block_idx > 0:
                    try:
                        ws.HPageBreaks.Add(ws.Rows(start_row))
                    except Exception:
                        pass
            try:
                last_row = max(1, len(chunk_items) * block_height)
                ws.PageSetup.PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(last_row, block_last_col)).Address
            except Exception:
                pass

        if created_sheets:
            try:
                template_sheet.Visible = 0  # xlSheetHidden
            except Exception:
                pass
        else:
            template_sheet.Name = "Этикетки"

        dst_wb.SaveAs(os.path.abspath(temp_output), FileFormat=56)
        os.makedirs(os.path.dirname(os.path.abspath(save_path)), exist_ok=True)
        shutil.copyfile(temp_output, save_path)
        return 0
    finally:
        try:
            if src_wb is not None:
                src_wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if dst_wb is not None:
                dst_wb.Close(SaveChanges=False)
        except Exception:
            pass
        if temp_dir and os.path.isdir(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)


def _run_preview(payload: dict, excel) -> int:
    xls_path = os.path.abspath(payload["xls_path"])
    wb = None
    try:
        excel.Visible = True
        excel.ScreenUpdating = True
        wb = excel.Workbooks.Open(
            xls_path,
            UpdateLinks=False,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
            Notify=False,
            AddToMru=False,
        )
        # Держим процесс пока пользователь не закроет workbook.
        while True:
            time.sleep(0.3)
            try:
                _ = wb.Name
            except Exception:
                break
        return 0
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass


def _run_print(payload: dict, excel) -> int:
    xls_path = os.path.abspath(payload["xls_path"])
    margins = payload.get("margins") or {}
    top_cm = float(margins.get("top_cm", 2.0))
    right_cm = float(margins.get("right_cm", 2.0))
    bottom_cm = float(margins.get("bottom_cm", 0.0))
    left_cm = float(margins.get("left_cm", 0.0))
    requested_printer = (payload.get("printer_name") or "").strip()

    wb = None
    try:
        wb = excel.Workbooks.Open(
            xls_path,
            UpdateLinks=False,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
            Notify=False,
            AddToMru=False,
        )
        if requested_printer:
            try:
                excel.ActivePrinter = requested_printer
            except Exception:
                pass
        used_printer = ""
        try:
            used_printer = str(excel.ActivePrinter or "")
        except Exception:
            used_printer = requested_printer

        for ws in wb.Worksheets:
            try:
                ps = ws.PageSetup
                ps.TopMargin = excel.CentimetersToPoints(top_cm)
                ps.RightMargin = excel.CentimetersToPoints(right_cm)
                ps.BottomMargin = excel.CentimetersToPoints(bottom_cm)
                ps.LeftMargin = excel.CentimetersToPoints(left_cm)
            except Exception:
                continue
        wb.PrintOut()
        print(json.dumps({"used_printer": used_printer}, ensure_ascii=False))
        return 0
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass


def _run_printers() -> int:
    try:
        import win32print
        result: list[str] = []
        for flags in (2, 4):
            for p in win32print.EnumPrinters(flags):
                name = p[2]
                if name and name not in result:
                    result.append(name)
        print(json.dumps({"printers": result}, ensure_ascii=False))
        return 0
    except Exception:
        print(json.dumps({"printers": []}, ensure_ascii=False))
        return 0


def main() -> int:
    if len(sys.argv) < 2:
        print("Missing payload path", file=sys.stderr)
        return 2
    payload_path = sys.argv[1]
    with open(payload_path, "r", encoding="utf-8") as fh:
        payload = json.load(fh)
    mode = str(payload.get("mode") or "generate").lower()

    if mode == "printers":
        return _run_printers()

    import pythoncom
    import win32com.client as win32
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        try:
            excel.Visible = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False
            excel.DisplayAlerts = False
        except Exception:
            pass
        if mode == "preview":
            return _run_preview(payload, excel)
        if mode == "print":
            return _run_print(payload, excel)
        return _run_generate(payload, excel)
    finally:
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
