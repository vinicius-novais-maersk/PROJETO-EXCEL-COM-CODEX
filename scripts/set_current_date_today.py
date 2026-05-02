from __future__ import annotations

import shutil
import sys
import time
from datetime import datetime
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client

try:
    from scripts.workbook_paths import resolve_workbook_path
except ModuleNotFoundError:
    from workbook_paths import resolve_workbook_path

WORKBOOK_PATH = resolve_workbook_path()
BACKUP_DIR = Path(__file__).resolve().parents[1] / "backups"


def log(message: str) -> None:
    print(f"[set-today] {message}", flush=True)


def backup_workbook(workbook_path: Path) -> Path:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = BACKUP_DIR / f"{workbook_path.stem}_backup_{timestamp}{workbook_path.suffix}"
    shutil.copy2(workbook_path, backup_path)
    return backup_path


def call_with_retry(func, *args, retries: int = 30, delay: float = 0.3, **kwargs):
    last_exc = None
    for _ in range(retries):
        try:
            return func(*args, **kwargs)
        except pywintypes.com_error as exc:
            last_exc = exc
            time.sleep(delay)
            pythoncom.PumpWaitingMessages()
    raise last_exc


def set_formula(worksheet, cell: str, formula: str) -> None:
    call_with_retry(setattr, worksheet.Range(cell), "Formula", formula)


def set_value(worksheet, cell: str, value) -> None:
    call_with_retry(setattr, worksheet.Range(cell), "Value", value)


def main() -> int:
    if not WORKBOOK_PATH.exists():
        log(f"Workbook not found: {WORKBOOK_PATH}")
        return 1

    backup_path = backup_workbook(WORKBOOK_PATH)
    log(f"Backup created at: {backup_path}")

    pythoncom.CoInitialize()
    excel = None
    workbook = None

    try:
        excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        call_with_retry(setattr, excel, "Visible", False)
        call_with_retry(setattr, excel, "DisplayAlerts", False)
        call_with_retry(setattr, excel, "EnableEvents", False)
        workbook = call_with_retry(
            excel.Workbooks.Open,
            str(WORKBOOK_PATH),
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
            Notify=False,
        )

        call_with_retry(setattr, excel, "Calculation", -4135)
        call_with_retry(setattr, excel, "CalculateBeforeSave", False)

        for sheet_name in ["Volume_DS", "Volume_MAO", "Volume_Graph"]:
            ws = workbook.Worksheets(sheet_name)
            set_value(ws, "L1", "Today")
            set_formula(ws, "M1", "=TODAY()")
            log(f"Updated header date in {sheet_name}")

        set_formula(workbook.Worksheets("Week_Overview"), "AH1", "=TODAY()")
        set_formula(workbook.Worksheets("Month2date"), "R1", "=TODAY()")
        set_formula(workbook.Worksheets("LastMonth_Overview"), "R1", "=TODAY()")
        log("Updated workbook date anchors to TODAY()")

        call_with_retry(workbook.Save)
        log("Workbook saved successfully.")
        return 0
    finally:
        if workbook is not None:
            try:
                call_with_retry(workbook.Close, SaveChanges=True)
            except Exception:
                pass
        if excel is not None:
            try:
                call_with_retry(excel.Quit)
            except Exception:
                pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    sys.exit(main())
