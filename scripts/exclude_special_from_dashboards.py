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
TARGET_SHEETS = ["Volume_MAO", "Volume_Graph", "Volume_DS", "Week_Overview"]
SPECIAL_RANGE = "'ROE_wk'!$BO:$BO"
SPECIAL_CRITERIA = '"<>Especial"'


def log(message: str) -> None:
    print(f"[exclude-special] {message}", flush=True)


def backup_workbook(workbook_path: Path) -> Path:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = BACKUP_DIR / f"{workbook_path.stem}_backup_{timestamp}{workbook_path.suffix}"
    shutil.copy2(workbook_path, backup_path)
    return backup_path


def call_with_retry(func, *args, retries: int = 30, delay: float = 0.4, **kwargs):
    last_exc = None
    for _ in range(retries):
        try:
            return func(*args, **kwargs)
        except pywintypes.com_error as exc:
            last_exc = exc
            time.sleep(delay)
            pythoncom.PumpWaitingMessages()
    raise last_exc


def find_matching_paren(text: str, open_index: int) -> int:
    depth = 0
    for index in range(open_index, len(text)):
        char = text[index]
        if char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
            if depth == 0:
                return index
    raise ValueError(f"Unbalanced parentheses in formula: {text}")


def split_top_level_args(text: str) -> list[str]:
    args: list[str] = []
    depth = 0
    start = 0
    for index, char in enumerate(text):
        if char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
        elif char == ";" and depth == 0:
            args.append(text[start:index])
            start = index + 1
    args.append(text[start:])
    return args


def transform_function(function_name: str, inner_text: str) -> str:
    args = [transform_formula_fragment(arg) for arg in split_top_level_args(inner_text)]
    joined = ";".join(args)

    if function_name == "CONT.SES":
        if "'ROE_wk'!" in joined and SPECIAL_RANGE not in joined:
            args.extend([SPECIAL_RANGE, SPECIAL_CRITERIA])
        return ";".join(args)

    if function_name == "FILTRO":
        if len(args) >= 2 and "'ROE_wk'!" in args[0] and SPECIAL_RANGE not in args[1]:
            args[1] = f"({args[1]}) * ({SPECIAL_RANGE}<>\"Especial\")"
        return ";".join(args)

    return joined


def transform_formula_fragment(text: str) -> str:
    target_functions = ("CONT.SES", "FILTRO")
    output: list[str] = []
    index = 0

    while index < len(text):
        matched_name = next(
            (name for name in target_functions if text.startswith(f"{name}(", index)),
            None,
        )
        if not matched_name:
            output.append(text[index])
            index += 1
            continue

        open_index = index + len(matched_name)
        close_index = find_matching_paren(text, open_index)
        inner_text = text[open_index + 1 : close_index]
        transformed_inner = transform_function(matched_name, inner_text)
        output.append(f"{matched_name}({transformed_inner})")
        index = close_index + 1

    return "".join(output)


def should_transform(formula: str) -> bool:
    return (
        isinstance(formula, str)
        and formula.startswith("=")
        and "'ROE_wk'!" in formula
        and SPECIAL_RANGE not in formula
    )


def update_sheet_formulas(worksheet) -> int:
    used = worksheet.UsedRange
    rows = used.Rows.Count
    cols = used.Columns.Count
    target_range = worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(rows, cols))
    formulas = call_with_retry(lambda: target_range.FormulaLocal)

    changed_cells = 0
    updated_rows: list[tuple] = []
    for row_index, row in enumerate(formulas, start=1):
        updated_row = list(row)
        for col_index, formula in enumerate(row, start=1):
            if not should_transform(formula):
                continue

            updated_formula = transform_formula_fragment(formula)
            if updated_formula == formula:
                continue

            updated_row[col_index - 1] = updated_formula
            changed_cells += 1
        updated_rows.append(tuple(updated_row))

    if changed_cells > 0:
        call_with_retry(setattr, target_range, "FormulaLocal", tuple(updated_rows))

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

        summary: dict[str, int] = {}
        for sheet_name in TARGET_SHEETS:
            log(f"Updating {sheet_name}...")
            worksheet = workbook.Worksheets(sheet_name)
            summary[sheet_name] = update_sheet_formulas(worksheet)

        call_with_retry(setattr, excel, "Calculation", -4105)
        call_with_retry(excel.CalculateFullRebuild)
        call_with_retry(workbook.Save)
        log(f"Workbook saved successfully. Changes: {summary}")

        validation = {
            "Volume_Graph_F12": call_with_retry(
                lambda: workbook.Worksheets("Volume_Graph").Range("F12").FormulaLocal
            ),
            "Volume_DS_F11": call_with_retry(
                lambda: workbook.Worksheets("Volume_DS").Range("F11").FormulaLocal
            ),
            "Week_Overview_D13": call_with_retry(
                lambda: workbook.Worksheets("Week_Overview").Range("D13").FormulaLocal
            ),
            "Volume_MAO_O5": call_with_retry(
                lambda: workbook.Worksheets("Volume_MAO").Range("O5").FormulaLocal
            ),
        }
        log(f"Validation: {validation}")

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
