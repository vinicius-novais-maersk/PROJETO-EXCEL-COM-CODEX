from __future__ import annotations

import shutil
import sys
import time
from datetime import datetime
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from scripts.exclude_special_from_dashboards import (
    WORKBOOK_PATH,
    backup_workbook,
    call_with_retry,
    should_transform,
    transform_formula_fragment,
)


def log(message: str) -> None:
    print(f"[week-overview-special] {message}", flush=True)


def update_week_overview(worksheet) -> int:
    used = worksheet.UsedRange
    rows = used.Rows.Count
    cols = used.Columns.Count
    formulas = call_with_retry(
        lambda: worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(rows, cols)).FormulaLocal
    )

    changed_cells = 0
    for row_index, row in enumerate(formulas, start=1):
        for col_index, formula in enumerate(row, start=1):
            if not should_transform(formula):
                continue

            updated_formula = transform_formula_fragment(formula)
            if updated_formula == formula:
                continue

            call_with_retry(
                setattr,
                worksheet.Cells(row_index, col_index),
                "FormulaLocal",
                updated_formula,
            )
            changed_cells += 1

    return changed_cells


def main() -> int:
    if not WORKBOOK_PATH.exists():
        log(f"Workbook not found: {WORKBOOK_PATH}")
        return 1

    backup_path = backup_workbook(WORKBOOK_PATH)
    log(f"Backup created at: {backup_path}")

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False

    workbook = None
    try:
        workbook = call_with_retry(
            excel.Workbooks.Open,
            str(WORKBOOK_PATH),
            UpdateLinks=False,
            ReadOnly=False,
        )
        worksheet = workbook.Worksheets("Week_Overview")
        changes = update_week_overview(worksheet)
        call_with_retry(setattr, excel, "Calculation", -4105)
        call_with_retry(excel.CalculateFullRebuild)
        call_with_retry(workbook.Save)
        log(f"Workbook saved successfully. Changes: {changes}")
        log(f"D13 -> {call_with_retry(lambda: worksheet.Range('D13').FormulaLocal)}")
        log(f"N13 -> {call_with_retry(lambda: worksheet.Range('N13').FormulaLocal)}")
        call_with_retry(workbook.Close, SaveChanges=False)
        workbook = None
        return 0
    except Exception as exc:  # noqa: BLE001
        log(f"Failed to update workbook: {exc}")
        if workbook is not None:
            try:
                call_with_retry(workbook.Close, SaveChanges=False)
            except Exception:  # noqa: BLE001
                pass
        return 1
    finally:
        excel.Quit()


if __name__ == "__main__":
    sys.exit(main())
