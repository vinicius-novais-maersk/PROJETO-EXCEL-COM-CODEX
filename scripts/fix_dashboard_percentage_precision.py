from __future__ import annotations

from pathlib import Path

import pythoncom
import win32com.client

try:
    from scripts.workbook_paths import resolve_workbook_path
except ModuleNotFoundError:
    from workbook_paths import resolve_workbook_path

WORKBOOK_PATH = resolve_workbook_path()


def remove_round_wrapper(formula: str) -> str:
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula
    marker = "=IFERROR(ROUND("
    if formula.startswith(marker) and formula.endswith(',2),"")'):
        inner = formula[len(marker) : -len(',2),"")')]
        return '=IFERROR(' + inner + ',"")'
    if formula.startswith("=ROUND(") and formula.endswith(",2)"):
        inner = formula[len("=ROUND(") : -len(",2)")]
        return "=" + inner
    return formula


def main() -> None:
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False

    workbook = None
    try:
        workbook = excel.Workbooks.Open(str(WORKBOOK_PATH), UpdateLinks=0, ReadOnly=False)

        # Week_Overview: keep the same logic, but stop rounding the total BR split KPIs.
        week = workbook.Worksheets("Week_Overview")
        for ref in ("N23", "P23"):
            week.Range(ref).Formula = remove_round_wrapper(week.Range(ref).Formula)

        # Daily OTO rows: remove formula-level rounding and rely on cell format only.
        for sheet_name, refs in {
            "Volume_DS": ("F16:M16",),
            "Volume_Graph": ("F17:M17",),
            "Volume_MAO": ("F19:M19",),
        }.items():
            ws = workbook.Worksheets(sheet_name)
            for ref in refs:
                rng = ws.Range(ref)
                formulas = rng.Formula
                updated = []
                for row in formulas:
                    updated.append(tuple(remove_round_wrapper(cell_formula) for cell_formula in row))
                rng.Formula = tuple(updated)

        # Standardize visible percentage display to one decimal place in the graph sheets.
        for sheet_name, ranges in {
            "Volume_DS": ("F16:M17", "I25:J27", "M25:N27", "Q25:R27"),
            "Volume_Graph": ("F17:M18", "I25:J27", "M25:N27", "Q25:R27"),
            "Volume_MAO": ("F19:M20", "I27:J29", "M27:N29", "Q27:R29"),
        }.items():
            ws = workbook.Worksheets(sheet_name)
            for ref in ranges:
                ws.Range(ref).NumberFormatLocal = "0,0%"

        workbook.Save()
        print("Dashboard percentage precision standardized.")
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
