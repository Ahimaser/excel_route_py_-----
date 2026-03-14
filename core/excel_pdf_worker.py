from __future__ import annotations

import os
import sys


def main() -> int:
    if len(sys.argv) < 3:
        print("Usage: excel_pdf_worker.py <xls_path> <pdf_path>", file=sys.stderr)
        return 2

    xls_path = os.path.abspath(sys.argv[1])
    pdf_path = os.path.abspath(sys.argv[2])
    os.makedirs(os.path.dirname(pdf_path), exist_ok=True)

    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        try:
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False
        except Exception:
            pass

        wb = excel.Workbooks.Open(
            xls_path,
            UpdateLinks=False,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
            Notify=False,
            AddToMru=False,
        )
        # 0 = xlTypePDF
        wb.ExportAsFixedFormat(0, pdf_path)
        return 0
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
