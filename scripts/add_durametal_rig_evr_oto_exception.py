from __future__ import annotations

import sys
from pathlib import Path

import pythoncom
import win32com.client

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from scripts.update_special_clients import WORKBOOK_PATH, backup_workbook, call_with_retry


SOURCE_NAME = "fDURAMETAL_RIG_EVR_CAB_OTO"
MATCH_TYPE = "Embarcador+Provedor+Porto+Brand"
SHIPPER = "DURAMETAL LTDA"
PORT = "Rio Grande"
BRAND = "Aliança"
PROVIDERS = [
    "EXPRESSO VALE REAL",
    "EXPRESSO VALE REAL LTDA",
]


def log(message: str) -> None:
    print(f"[durametal-rig-evr-oto] {message}", flush=True)


def get_data_end_row(roe_table) -> int:
    return roe_table.Range.Row + roe_table.Range.Rows.Count - 1


def patch_formula(formula: str) -> str:
    if (
        MATCH_TYPE in formula
        and 'CONT.SES(Exceptions_Mapping[ExceptionType]; tipo; Exceptions_Mapping[MatchType]; "Embarcador+Provedor+Porto+Brand"'
        in formula
    ):
        return formula

    formula = formula.replace(
        "embcliprovporto; CC2;\n    emblocal;",
        "embcliprovporto; CC2;\n    embprovportobrand; K2&D2&AY2&AI2;\n    emblocal;",
    )
    formula = formula.replace(
        'vetorEmbCliProvPorto; SE((tp=tipo)*(mt="Embarcador+Cliente Proposta+Provedor+Porto"); mv; "");\n    vetorEmbLocal;',
        'vetorEmbCliProvPorto; SE((tp=tipo)*(mt="Embarcador+Cliente Proposta+Provedor+Porto"); mv; "");\n    vetorEmbProvPortoBrand; SE((tp=tipo)*(mt="Embarcador+Provedor+Porto+Brand"); mv; "");\n    vetorEmbLocal;',
    )
    formula = formula.replace(
        "E(embcliprovporto<>\"\"; ÉNÚM(CORRESPX(embcliprovporto; vetorEmbCliProvPorto)));\n        E(emblocal<>\"\";",
        "E(embcliprovporto<>\"\"; ÉNÚM(CORRESPX(embcliprovporto; vetorEmbCliProvPorto)));\n        E(embprovportobrand<>\"\"; CONT.SES(Exceptions_Mapping[ExceptionType]; tipo; Exceptions_Mapping[MatchType]; \"Embarcador+Provedor+Porto+Brand\"; Exceptions_Mapping[MatchValue]; embprovportobrand)>0);\n        E(emblocal<>\"\";",
    )
    formula = formula.replace(
        "E(embprovportobrand<>\"\"; ÉNÚM(CORRESPX(embprovportobrand; vetorEmbProvPortoBrand)));",
        "E(embprovportobrand<>\"\"; CONT.SES(Exceptions_Mapping[ExceptionType]; tipo; Exceptions_Mapping[MatchType]; \"Embarcador+Provedor+Porto+Brand\"; Exceptions_Mapping[MatchValue]; embprovportobrand)>0);",
    )
    return formula


def ensure_formula_support(roe_ws, roe_table) -> None:
    data_end_row = get_data_end_row(roe_table)
    for column in ("CA", "CB"):
        first_cell = roe_ws.Range(f"{column}2")
        patched_formula = patch_formula(first_cell.FormulaLocal)
        if MATCH_TYPE not in patched_formula:
            raise RuntimeError(f"Could not patch {column} formula with {MATCH_TYPE}.")
        call_with_retry(
            setattr,
            roe_ws.Range(f"{column}2:{column}{data_end_row}"),
            "FormulaLocal",
            patched_formula,
        )


def ensure_mapping_rows(mapping_table) -> int:
    existing = set()
    data_rows = mapping_table.DataBodyRange.Rows.Count if mapping_table.DataBodyRange is not None else 0
    if data_rows:
        for index in range(1, data_rows + 1):
            row_range = mapping_table.DataBodyRange.Rows(index)
            key = (
                row_range.Cells(1, 1).Value,
                row_range.Cells(1, 2).Value,
                row_range.Cells(1, 3).Value,
                row_range.Cells(1, 4).Value,
            )
            existing.add(key)

    added = 0
    for provider in PROVIDERS:
        match_value = f"{SHIPPER}{provider}{PORT}{BRAND}"
        key = ("OTO", SOURCE_NAME, MATCH_TYPE, match_value)
        if key in existing:
            continue
        new_row = mapping_table.ListRows.Add()
        row_range = new_row.Range
        call_with_retry(setattr, row_range.Cells(1, 1), "Value", "OTO")
        call_with_retry(setattr, row_range.Cells(1, 2), "Value", SOURCE_NAME)
        call_with_retry(setattr, row_range.Cells(1, 3), "Value", MATCH_TYPE)
        call_with_retry(setattr, row_range.Cells(1, 4), "Value", match_value)
        existing.add(key)
        added += 1
    return added


def main() -> int:
    if not WORKBOOK_PATH.exists():
        log(f"Workbook not found: {WORKBOOK_PATH}")
        return 1

    backup_path = backup_workbook(WORKBOOK_PATH)
    log(f"Backup created at: {backup_path}")

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    excel.ScreenUpdating = False

    workbook = None
    try:
        workbook = call_with_retry(
            excel.Workbooks.Open,
            str(WORKBOOK_PATH),
            UpdateLinks=False,
            ReadOnly=False,
        )
        roe_ws = workbook.Worksheets.Item("ROE_wk")
        mapping_ws = workbook.Worksheets.Item("Exceptions_Mapping")
        roe_table = roe_ws.ListObjects.Item("ROE_wk")
        mapping_table = mapping_ws.ListObjects.Item("Exceptions_Mapping")

        ensure_formula_support(roe_ws, roe_table)
        added = ensure_mapping_rows(mapping_table)

        # Recalculate only the affected row pattern and dependent OTO Out column.
        call_with_retry(roe_ws.Range("CA:CB").Calculate)
        call_with_retry(roe_ws.Range("BL:BL").Calculate)

        call_with_retry(workbook.Save)
        log(f"Added mapping rows: {added}")

        call_with_retry(workbook.Close, SaveChanges=False, retries=120, delay=1.0)
        workbook = None
        return 0
    except Exception as exc:  # noqa: BLE001
        log(f"Failed to update workbook: {exc}")
        if workbook is not None:
            try:
                workbook.Close(False)
            except Exception:  # noqa: BLE001
                pass
        return 1
    finally:
        try:
            excel.Quit()
        except Exception:  # noqa: BLE001
            pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
