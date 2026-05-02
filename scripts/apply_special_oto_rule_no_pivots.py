from __future__ import annotations

import sys
from pathlib import Path

import pythoncom
import win32com.client

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from scripts.fix_oto_special_rule_complete import (
    update_roe_late_flag,
    update_volume_ds,
    update_volume_graph,
    update_volume_mao,
    update_week_sheet,
)
from scripts.update_special_clients import (
    WORKBOOK_PATH,
    backup_workbook,
    call_with_retry,
)


def log(message: str) -> None:
    print(f"[special-oto-no-pivots] {message}", flush=True)


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
        call_with_retry(setattr, excel, "ScreenUpdating", False)

        workbook = call_with_retry(
            excel.Workbooks.Open,
            str(WORKBOOK_PATH),
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
            Notify=False,
        )

        changed: dict[str, int] = {}
        update_roe_late_flag(workbook.Worksheets("ROE_wk"))
        changed["ROE_wk[Atrasado?]"] = 1
        changed["Week_Overview"] = update_week_sheet(
            workbook.Worksheets("Week_Overview"), "ROE_wk"
        )
        changed["Volume_DS"] = update_volume_ds(workbook.Worksheets("Volume_DS"))
        changed["Volume_Graph"] = update_volume_graph(
            workbook.Worksheets("Volume_Graph")
        )
        changed["Volume_MAO"] = update_volume_mao(workbook.Worksheets("Volume_MAO"))

        call_with_retry(setattr, excel, "Calculation", -4105)
        call_with_retry(excel.CalculateFullRebuild)
        call_with_retry(workbook.Save)

        checks = {
            "ROE_wk!Atrasado?_formula": call_with_retry(
                lambda: (
                    workbook.Worksheets("ROE_wk")
                    .ListObjects("ROE_wk")
                    .ListColumns("Atrasado?")
                    .DataBodyRange.Cells(1, 1)
                    .FormulaLocal
                )
            ),
            "Week_Overview!Q23_formula": call_with_retry(
                lambda: workbook.Worksheets("Week_Overview").Range("Q23").FormulaLocal
            ),
            "Volume_DS!M16_formula": call_with_retry(
                lambda: workbook.Worksheets("Volume_DS").Range("M16").FormulaLocal
            ),
            "Volume_Graph!M17_formula": call_with_retry(
                lambda: workbook.Worksheets("Volume_Graph").Range("M17").FormulaLocal
            ),
            "Volume_MAO!M19_formula": call_with_retry(
                lambda: workbook.Worksheets("Volume_MAO").Range("M19").FormulaLocal
            ),
        }
        log(f"Changes applied: {changed}")
        log(f"Checks: {checks}")
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
