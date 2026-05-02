from __future__ import annotations

import os
import sys
from pathlib import Path

import pythoncom
import win32com.client

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from scripts.update_special_clients import WORKBOOK_PATH, call_with_retry, update_special_formula


def main() -> int:
    workbook_path = Path(os.environ.get("DSU_WORKBOOK_PATH", str(WORKBOOK_PATH)))
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    workbook = None
    try:
        workbook = call_with_retry(
            excel.Workbooks.Open,
            str(workbook_path),
            UpdateLinks=False,
            ReadOnly=False,
        )
        roe_ws = workbook.Worksheets("ROE_wk")
        try:
            update_special_formula(roe_ws)
            print("FORMULA_OK")
        except Exception as exc:  # noqa: BLE001
            print(f"FORMULA_FAIL: {exc!r}")
            sample = roe_ws.ListObjects("ROE_wk").ListColumns("Especiais").DataBodyRange.Cells(1, 1)
            print("CURRENT_FORMULA_LOCAL:")
            print(sample.FormulaLocal)
            return 1
        call_with_retry(workbook.Close, SaveChanges=False)
        workbook = None
        return 0
    finally:
        if workbook is not None:
            try:
                call_with_retry(workbook.Close, SaveChanges=False)
            except Exception:
                pass
        excel.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
