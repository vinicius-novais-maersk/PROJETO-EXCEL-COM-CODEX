from __future__ import annotations

from pathlib import Path

import pythoncom
import win32com.client

try:
    from scripts.workbook_paths import resolve_workbook_path
except ModuleNotFoundError:
    from workbook_paths import resolve_workbook_path

WORKBOOK_PATH = resolve_workbook_path()
ROE_SHEET = "ROE_wk"
HEADER_CELL = "CJ1"
FIRST_DATA_ROW = 2
HEADER = "Origem_Status_SIL"
FORMULA = (
    '=IF(XLOOKUP(A2,SIL_wk!B:B,SIL_wk!B:B,"")="",'
    '"OS nao encontrada no SIL",'
    'IF(XLOOKUP(A2,SIL_wk!Y:Y,SIL_wk!Y:Y,"")="Sem Preenchimento",'
    '"Sem preenchimento no SIL",'
    '"Status encontrado no SIL"))'
)


def main() -> None:
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False

    workbook = None
    try:
        workbook = excel.Workbooks.Open(str(WORKBOOK_PATH), UpdateLinks=0, ReadOnly=False)
        ws = workbook.Worksheets(ROE_SHEET)
        last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row

        ws.Range(HEADER_CELL).Value = HEADER
        ws.Range(f"CJ{FIRST_DATA_ROW}:CJ{last_row}").Formula = FORMULA
        ws.Range("CJ:CJ").EntireColumn.AutoFit()

        workbook.Save()
        print(f"Added {HEADER} to {ROE_SHEET}!CJ with {last_row - 1} formulas.")
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
