from __future__ import annotations

import sys
from pathlib import Path

import win32com.client

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from scripts.update_special_clients import WORKBOOK_PATH, backup_workbook, call_with_retry


TARGET_SHEET = "Volume_DS"
TARGET_COLUMNS = ["F", "G", "H", "I", "J", "K", "L"]
TARGET_ROW = 16
ALIANCA_FILTER = f"<>Alian{chr(231)}a"


def log(message: str) -> None:
    print(f"[volume-ds-let] {message}", flush=True)


def build_formula(day_column: str) -> str:
    day_ref = f"{day_column}$8"
    return (
        "=LET("
        "Base;CONT.SES("
        "'ROE_wk'!$AV:$AV;\"Ok\";"
        "'ROE_wk'!$AY:$AY;$M$3;"
        "'ROE_wk'!$AR:$AR;$M$2;"
        f"'ROE_wk'!$AP:$AP;{day_ref};"
        f"'ROE_wk'!$AI:$AI;\"{ALIANCA_FILTER}\";"
        "'ROE_wk'!$BB:$BB;\"<>Sem Preenchimento\";"
        "'ROE_wk'!$BL:$BL;\"N\";"
        "'ROE_wk'!$BO:$BO;\"<>Especial\""
        ");"
        "Atrasados;CONT.SES("
        "'ROE_wk'!$AV:$AV;\"Ok\";"
        "'ROE_wk'!$AY:$AY;$M$3;"
        "'ROE_wk'!$AR:$AR;$M$2;"
        f"'ROE_wk'!$AP:$AP;{day_ref};"
        f"'ROE_wk'!$AI:$AI;\"{ALIANCA_FILTER}\";"
        "'ROE_wk'!$BB:$BB;\"Atrasado\";"
        "'ROE_wk'!$BL:$BL;\"N\";"
        "'ROE_wk'!$BO:$BO;\"<>Especial\""
        ");"
        "SE(Base=0;\"\";ARRED(1-Atrasados/Base;2))"
        ")"
    )


def update_formulas(worksheet) -> None:
    for column_index, column in enumerate(TARGET_COLUMNS, start=6):
        cell = worksheet.Cells(TARGET_ROW, column_index)
        formula = build_formula(column)
        call_with_retry(setattr, cell, "FormulaLocal", formula)


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
        worksheet = workbook.Worksheets(TARGET_SHEET)
        update_formulas(worksheet)

        call_with_retry(setattr, excel, "Calculation", -4105)
        call_with_retry(excel.CalculateFullRebuild)
        call_with_retry(workbook.Save)

        for cell_ref, column_index in {"F16": 6, "J16": 10, "L16": 12}.items():
            formula = call_with_retry(lambda idx=column_index: worksheet.Cells(TARGET_ROW, idx).FormulaLocal)
            value = call_with_retry(lambda idx=column_index: worksheet.Cells(TARGET_ROW, idx).Value)
            log(f"{cell_ref} formula -> {formula}")
            log(f"{cell_ref} value -> {value}")

        call_with_retry(workbook.Close, SaveChanges=False)
        workbook = None
        log("Workbook saved successfully.")
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
        try:
            excel.Quit()
        except Exception:  # noqa: BLE001
            pass


if __name__ == "__main__":
    sys.exit(main())
