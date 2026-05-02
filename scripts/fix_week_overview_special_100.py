from __future__ import annotations

import os
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
PORT_ROWS = (2, 3, 6, 7, 8, 11, 12, 13, 16, 17, 18, 19, 20)
REGION_ROWS = (4, 9, 14, 21)
OTO_ROWS = PORT_ROWS + REGION_ROWS + (23,)


def log(message: str) -> None:
    print(f"[week-overview-special-100] {message}", flush=True)


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


def set_formula(ws, cell_ref: str, formula: str) -> None:
    call_with_retry(setattr, ws.Range(cell_ref), "Formula", formula)


def countifs_formula(*pairs: str) -> str:
    joined = ",".join(f"'{pair}'" if False else pair for pair in pairs)
    return f"COUNTIFS({joined})"


def build_scope(row: int) -> list[str]:
    if row in PORT_ROWS:
        return ["'ROE_wk'!$AY:$AY", f"$AG{row}"]
    if row in REGION_ROWS:
        return ["'ROE_wk'!$AZ:$AZ", f"$A{row}"]
    return []


def build_oto_formula(row: int, cost_center: str | None, round_result: bool = False) -> str:
    denominator_args = [
        "'ROE_wk'!$AV:$AV",
        '"Ok"',
        *build_scope(row),
        "'ROE_wk'!$AR:$AR",
        "$AG$1",
    ]
    numerator_args = denominator_args.copy()

    if cost_center is not None:
        denominator_args.extend(["'ROE_wk'!$AI:$AI", f'"{cost_center}"'])
        numerator_args.extend(["'ROE_wk'!$AI:$AI", f'"{cost_center}"'])

    denominator_args.extend(
        [
            "'ROE_wk'!$BB:$BB",
            '"<>Sem Preenchimento"',
            "'ROE_wk'!$BL:$BL",
            '"N"',
        ]
    )
    numerator_args.extend(
        [
            "'ROE_wk'!$BB:$BB",
            '"Atrasado"',
            "'ROE_wk'!$BL:$BL",
            '"N"',
            "'ROE_wk'!$BO:$BO",
            '"<>Especial"',
        ]
    )

    denominator = countifs_formula(*denominator_args)
    numerator = countifs_formula(*numerator_args)
    core = f'IF({denominator}=0,"",1-IFERROR({numerator}/{denominator},0))'
    if round_result:
        core = f"ROUND({core},2)"
    return f"={core}"


def update_week_overview(ws) -> None:
    for row in PORT_ROWS + REGION_ROWS:
        set_formula(ws, f"N{row}", build_oto_formula(row, "Aliança"))
        set_formula(ws, f"P{row}", build_oto_formula(row, "<>Aliança"))
        set_formula(ws, f"Q{row}", build_oto_formula(row, None))

    set_formula(ws, "N23", build_oto_formula(23, "Aliança", round_result=True))
    set_formula(ws, "P23", build_oto_formula(23, "<>Aliança", round_result=True))
    set_formula(ws, "Q23", build_oto_formula(23, None))

    set_formula(
        ws,
        "D23",
        '=COUNTIFS(\'ROE_wk\'!$AV:$AV,"Ok",\'ROE_wk\'!$AR:$AR,$AG$1,\'ROE_wk\'!$AI:$AI,"Aliança")',
    )
    set_formula(
        ws,
        "H23",
        '=COUNTIFS(\'ROE_wk\'!$AV:$AV,"Ok",\'ROE_wk\'!$AR:$AR,$AG$1,\'ROE_wk\'!$AI:$AI,"<>Aliança")',
    )
    set_formula(
        ws,
        "K23",
        '=COUNTIFS(\'ROE_wk\'!$AV:$AV,"Ok",\'ROE_wk\'!$AR:$AR,$AG$1)',
    )


def main() -> int:
    workbook_path = Path(os.environ.get("DSU_WORKBOOK_PATH", str(WORKBOOK_PATH)))
    if not workbook_path.exists():
        log(f"Workbook not found: {workbook_path}")
        return 1

    backup_path = backup_workbook(workbook_path)
    log(f"Backup created at: {backup_path}")

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
        call_with_retry(setattr, excel, "Calculation", -4135)

        ws = workbook.Worksheets("Week_Overview")
        update_week_overview(ws)

        call_with_retry(setattr, excel, "Calculation", -4105)
        call_with_retry(excel.CalculateFullRebuild)
        call_with_retry(workbook.Save)

        log(f"D23={call_with_retry(lambda: ws.Range('D23').Value)}")
        log(f"H23={call_with_retry(lambda: ws.Range('H23').Value)}")
        log(f"K23={call_with_retry(lambda: ws.Range('K23').Value)}")
        log(f"N23={call_with_retry(lambda: ws.Range('N23').Value)}")
        log(f"P23={call_with_retry(lambda: ws.Range('P23').Value)}")
        log(f"Q23={call_with_retry(lambda: ws.Range('Q23').Value)}")

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
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    sys.exit(main())
