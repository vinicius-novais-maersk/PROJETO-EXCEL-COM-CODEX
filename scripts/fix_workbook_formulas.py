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
    print(f"[fix-workbook] {message}", flush=True)


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


def set_formula_local(worksheet, cell: str, formula: str) -> None:
    call_with_retry(setattr, worksheet.Range(cell), "FormulaLocal", formula)


def cab_day_total(day_ref: str) -> str:
    return (
        "CONT.SES("
        "'ROE_wk'!$AV:$AV;\"Ok\";"
        "'ROE_wk'!$AY:$AY;$M$3;"
        "'ROE_wk'!$AR:$AR;$M$2;"
        f"'ROE_wk'!$AP:$AP;{day_ref};"
        "'ROE_wk'!$AI:$AI;\"Aliança\")"
    )


def cab_day_delayed(day_ref: str) -> str:
    return (
        "CONT.SES("
        "'ROE_wk'!$AV:$AV;\"Ok\";"
        "'ROE_wk'!$AY:$AY;$M$3;"
        "'ROE_wk'!$AR:$AR;$M$2;"
        f"'ROE_wk'!$AP:$AP;{day_ref};"
        "'ROE_wk'!$AI:$AI;\"Aliança\";"
        "'ROE_wk'!$BB:$BB;\"Atrasado\")"
    )


def cab_day_48h(day_ref: str) -> str:
    return (
        "CONT.SES("
        "'ROE_wk'!$AV:$AV;\"Ok\";"
        "'ROE_wk'!$AY:$AY;$M$3;"
        "'ROE_wk'!$AR:$AR;$M$2;"
        f"'ROE_wk'!$AP:$AP;{day_ref};"
        "'ROE_wk'!$AI:$AI;\"Aliança\";"
        "'ROE_wk'!$AT:$AT;\"N\")"
    )


def ds_day_total(day_ref: str) -> str:
    return (
        "CONT.SES("
        "'ROE_wk'!$AV:$AV;\"Ok\";"
        "'ROE_wk'!$AY:$AY;$M$3;"
        "'ROE_wk'!$AR:$AR;$M$2;"
        f"'ROE_wk'!$AP:$AP;{day_ref};"
        "'ROE_wk'!$AI:$AI;\"<>Aliança\")"
    )


def ds_day_denominator(day_ref: str) -> str:
    return (
        "CONT.SES("
        "'ROE_wk'!$AV:$AV;\"Ok\";"
        "'ROE_wk'!$AY:$AY;$M$3;"
        "'ROE_wk'!$AR:$AR;$M$2;"
        f"'ROE_wk'!$AP:$AP;{day_ref};"
        "'ROE_wk'!$AI:$AI;\"<>Aliança\";"
        "'ROE_wk'!$BB:$BB;\"<>Sem Preenchimento\";"
        "'ROE_wk'!$BL:$BL;\"N\")"
    )


def ds_day_delayed(day_ref: str) -> str:
    return (
        "CONT.SES("
        "'ROE_wk'!$AV:$AV;\"Ok\";"
        "'ROE_wk'!$AY:$AY;$M$3;"
        "'ROE_wk'!$AR:$AR;$M$2;"
        f"'ROE_wk'!$AP:$AP;{day_ref};"
        "'ROE_wk'!$AI:$AI;\"<>Aliança\";"
        "'ROE_wk'!$BB:$BB;\"Atrasado\";"
        "'ROE_wk'!$BL:$BL;\"N\")"
    )


def ds_day_48h(day_ref: str) -> str:
    return (
        "CONT.SES("
        "'ROE_wk'!$AV:$AV;\"Ok\";"
        "'ROE_wk'!$AY:$AY;$M$3;"
        "'ROE_wk'!$AR:$AR;$M$2;"
        f"'ROE_wk'!$AP:$AP;{day_ref};"
        "'ROE_wk'!$AI:$AI;\"<>Aliança\";"
        "'ROE_wk'!$AT:$AT;\"N\";"
        "'ROE_wk'!$BL:$BL;\"N\")"
    )


def oto_cab_day_formula(day_col: str) -> str:
    day_ref = f"{day_col}$8"
    total = cab_day_total(day_ref)
    delayed = cab_day_delayed(day_ref)
    return f'=SEERRO(ARRED(SE({total}=0;"";1-{delayed}/{total});2);"")'


def schedule_48h_cab_day_formula(day_col: str) -> str:
    day_ref = f"{day_col}$8"
    total = cab_day_total(day_ref)
    delayed = cab_day_48h(day_ref)
    return f'=SE({total}=0;"";1-SEERRO({delayed}/{total};0))'


def oto_ds_day_formula(day_col: str) -> str:
    day_ref = f"{day_col}$8"
    denominator = ds_day_denominator(day_ref)
    delayed = ds_day_delayed(day_ref)
    return f'=SEERRO(ARRED(SE({denominator}=0;"";1-{delayed}/{denominator});2);"")'


def schedule_48h_ds_day_formula(day_col: str) -> str:
    day_ref = f"{day_col}$8"
    total = ds_day_total(day_ref)
    delayed = ds_day_48h(day_ref)
    return f'=SE({total}=0;"";1-SEERRO({delayed}/{total};0))'


def current_week_lookup_formula(result_column: str) -> str:
    return f'=PROCX($M$3;Week_Overview!AG:AG;Week_Overview!{result_column}:{result_column};"")'


def historical_lookup_formula(week_cell: str, lookup_column: str, result_column: str) -> str:
    return (
        f'=SEERRO(PROCX($M$3;'
        f'INDIRETO("Week_"&${week_cell}&"!{lookup_column}:{lookup_column}");'
        f'INDIRETO("Week_"&${week_cell}&"!{result_column}:{result_column}");'
        f'"");"")'
    )


def apply_volume_graph_fixes(workbook) -> int:
    ws = workbook.Worksheets("Volume_Graph")
    changed = 0

    for col in ["F", "G", "H", "I", "J", "K", "L"]:
        set_formula_local(ws, f"{col}17", oto_cab_day_formula(col))
        set_formula_local(ws, f"{col}18", schedule_48h_cab_day_formula(col))
        changed += 2

    set_formula_local(ws, "M17", current_week_lookup_formula("N"))
    set_formula_local(ws, "M18", current_week_lookup_formula("R"))
    changed += 2

    for row in [25, 26]:
        set_formula_local(ws, f"J{row}", historical_lookup_formula(f"E{row}", "AG", "N"))
        set_formula_local(ws, f"N{row}", historical_lookup_formula(f"E{row}", "AG", "P"))
        set_formula_local(ws, f"R{row}", historical_lookup_formula(f"E{row}", "AG", "Q"))
        changed += 3

    set_formula_local(ws, "J27", current_week_lookup_formula("N"))
    set_formula_local(ws, "N27", current_week_lookup_formula("P"))
    set_formula_local(ws, "R27", current_week_lookup_formula("Q"))
    changed += 3

    return changed


def apply_volume_mao_fixes(workbook) -> int:
    ws = workbook.Worksheets("Volume_MAO")
    changed = 0

    for col in ["F", "G", "H", "I", "J", "K", "L"]:
        set_formula_local(ws, f"{col}19", oto_cab_day_formula(col))
        set_formula_local(ws, f"{col}20", schedule_48h_cab_day_formula(col))
        changed += 2

    set_formula_local(ws, "M19", current_week_lookup_formula("N"))
    set_formula_local(ws, "M20", current_week_lookup_formula("R"))
    changed += 2

    for row in [27, 28]:
        set_formula_local(ws, f"J{row}", historical_lookup_formula(f"E{row}", "AE", "N"))
        set_formula_local(ws, f"N{row}", historical_lookup_formula(f"E{row}", "AE", "P"))
        set_formula_local(ws, f"R{row}", historical_lookup_formula(f"E{row}", "AE", "Q"))
        changed += 3

    set_formula_local(ws, "J29", current_week_lookup_formula("N"))
    set_formula_local(ws, "N29", current_week_lookup_formula("P"))
    set_formula_local(ws, "R29", current_week_lookup_formula("Q"))
    changed += 3

    return changed


def apply_volume_ds_fixes(workbook) -> int:
    ws = workbook.Worksheets("Volume_DS")
    changed = 0

    for col in ["F", "G", "H", "I", "J", "K", "L"]:
        set_formula_local(ws, f"{col}16", oto_ds_day_formula(col))
        set_formula_local(ws, f"{col}17", schedule_48h_ds_day_formula(col))
        changed += 2

    set_formula_local(ws, "M16", current_week_lookup_formula("P"))
    set_formula_local(
        ws,
        "M17",
        '=SE(CONT.SES(\'ROE_wk\'!$AV:$AV;"Ok";\'ROE_wk\'!$AY:$AY;$M$3;\'ROE_wk\'!$AR:$AR;$M$2;\'ROE_wk\'!$AI:$AI;"<>Aliança")=0;"";'
        '1-SEERRO(CONT.SES(\'ROE_wk\'!$AV:$AV;"Ok";\'ROE_wk\'!$AY:$AY;$M$3;\'ROE_wk\'!$AR:$AR;$M$2;\'ROE_wk\'!$AI:$AI;"<>Aliança";\'ROE_wk\'!$AT:$AT;"N";\'ROE_wk\'!$BL:$BL;"N")/'
        'CONT.SES(\'ROE_wk\'!$AV:$AV;"Ok";\'ROE_wk\'!$AY:$AY;$M$3;\'ROE_wk\'!$AR:$AR;$M$2;\'ROE_wk\'!$AI:$AI;"<>Aliança");0))',
    )
    changed += 2

    for row in [25, 26]:
        set_formula_local(ws, f"J{row}", historical_lookup_formula(f"E{row}", "AE", "N"))
        set_formula_local(ws, f"N{row}", historical_lookup_formula(f"E{row}", "AE", "P"))
        set_formula_local(ws, f"R{row}", historical_lookup_formula(f"E{row}", "AE", "Q"))
        changed += 3

    set_formula_local(ws, "J27", current_week_lookup_formula("N"))
    set_formula_local(ws, "N27", current_week_lookup_formula("P"))
    set_formula_local(ws, "R27", current_week_lookup_formula("Q"))
    changed += 3

    return changed


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

        changed = 0
        changed += apply_volume_ds_fixes(workbook)
        changed += apply_volume_mao_fixes(workbook)
        changed += apply_volume_graph_fixes(workbook)

        call_with_retry(workbook.Save)

        log(f"Workbook updated successfully. Formula changes applied: {changed}")
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
