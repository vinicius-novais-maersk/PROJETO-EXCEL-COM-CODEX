from __future__ import annotations

import json
import sys
from pathlib import Path

import pythoncom
import win32com.client

try:
    from scripts.update_special_clients import call_with_retry
except ModuleNotFoundError:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    from scripts.update_special_clients import call_with_retry

WORKBOOK = Path(r"C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX\analysis\Base_DSU2026_WK19_rules_hypercare_fix_20260508_173020.xlsm")
REPORT = Path(r"C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX\analysis\Base_DSU2026_WK19_rules_hypercare_fix_20260508_173020_changes.json")
ALLIANCE = "Alian" + chr(231) + "a"
PORT_ROWS = (2, 3, 6, 7, 8, 11, 12, 13, 16, 17, 18, 19)
REGION_SUMS = {4: (2, 3), 9: (6, 8), 14: (11, 13), 20: (16, 19)}
TOTAL_ROW = 22


def q(value: str) -> str:
    return f'"{value}"'


def set_formula(ws, cell: str, formula: str, changes: list[dict]) -> None:
    before = call_with_retry(lambda: ws.Range(cell).Formula)
    call_with_retry(setattr, ws.Range(cell), "Formula", formula)
    changes.append({"sheet": ws.Name, "cell": cell, "before": str(before), "after": formula})


def countifs(*parts: str) -> str:
    return "COUNTIFS(" + ",".join(parts) + ")"


def crit(range_ref: str, criterion: str) -> str:
    return f"{range_ref},{criterion}"


def oto_formula(sheet_name: str, row: int, *, scope_type: str, alliance_filter: str | None) -> str:
    # KPI OTO rule: Sem Preenchimento is excluded from KPI denominator; Especial stays in denominator but is not counted as late.
    if scope_type == "port":
        scope = [crit(f"'{sheet_name}'!$AY:$AY", f"$Q{row}")]
    elif scope_type == "region":
        scope = [crit(f"'{sheet_name}'!$AZ:$AZ", f"$A{row}")]
    elif scope_type == "total":
        scope = []
    else:
        raise ValueError(scope_type)

    parts = [
        crit(f"'{sheet_name}'!$AV:$AV", q("Ok")),
        *scope,
        crit(f"'{sheet_name}'!$AR:$AR", "$Q$1"),
        crit(f"'{sheet_name}'!$BB:$BB", q("<>Sem Preenchimento")),
        crit(f"'{sheet_name}'!$BL:$BL", q("N")),
    ]
    if alliance_filter is not None:
        parts.append(crit(f"'{sheet_name}'!$AI:$AI", q(alliance_filter)))
    denominator = countifs(*parts)

    late_parts = [p for p in parts]
    # Replace OTD criterion and keep Especial out of the late numerator only.
    late_parts = [p for p in late_parts if "$BB:$BB" not in p]
    late_parts.append(crit(f"'{sheet_name}'!$BB:$BB", q("Atrasado")))
    late_parts.append(crit(f"'{sheet_name}'!$BO:$BO", q("<>Especial")))
    late = countifs(*late_parts)
    return f'=IF({denominator}=0,"",1-IFERROR({late}/{denominator},0))'


def monthly_count_formula(row: int, pattern: str, *, port: bool = True) -> str:
    parts = [
        crit("'ROE_wk_monthly'!$AV:$AV", q("Ok")),
        crit("'ROE_wk_monthly'!$BG:$BG", "$Q$1"),
        crit("'ROE_wk_monthly'!$AI:$AI", q(ALLIANCE)),
        crit("'ROE_wk_monthly'!$N:$N", q(pattern)),
    ]
    if port:
        parts.insert(1, crit("'ROE_wk_monthly'!$AY:$AY", f"$Q{row}"))
    return "=" + countifs(*parts)


def update_month_sheet(ws, changes: list[dict]) -> None:
    for row in PORT_ROWS:
        set_formula(ws, f"K{row}", oto_formula("ROE_wk", row, scope_type="port", alliance_filter=ALLIANCE), changes)
        set_formula(ws, f"L{row}", oto_formula("ROE_wk", row, scope_type="port", alliance_filter=f"<>{ALLIANCE}"), changes)
        set_formula(ws, f"M{row}", oto_formula("ROE_wk", row, scope_type="port", alliance_filter=None), changes)
        set_formula(ws, f"O{row}", monthly_count_formula(row, "*Reagenda*", port=True), changes)
        set_formula(ws, f"P{row}", monthly_count_formula(row, "*HYPERCARE*", port=True), changes)

    for row, (start, end) in REGION_SUMS.items():
        set_formula(ws, f"K{row}", oto_formula("ROE_wk", row, scope_type="region", alliance_filter=ALLIANCE), changes)
        set_formula(ws, f"L{row}", oto_formula("ROE_wk", row, scope_type="region", alliance_filter=f"<>{ALLIANCE}"), changes)
        set_formula(ws, f"M{row}", oto_formula("ROE_wk", row, scope_type="region", alliance_filter=None), changes)
        set_formula(ws, f"O{row}", f"=SUM(O{start}:O{end})", changes)
        set_formula(ws, f"P{row}", f"=SUM(P{start}:P{end})", changes)

    set_formula(ws, f"K{TOTAL_ROW}", oto_formula("ROE_wk", TOTAL_ROW, scope_type="total", alliance_filter=ALLIANCE), changes)
    set_formula(ws, f"L{TOTAL_ROW}", oto_formula("ROE_wk", TOTAL_ROW, scope_type="total", alliance_filter=f"<>{ALLIANCE}"), changes)
    set_formula(ws, f"M{TOTAL_ROW}", oto_formula("ROE_wk", TOTAL_ROW, scope_type="total", alliance_filter=None), changes)
    set_formula(ws, f"O{TOTAL_ROW}", "=SUM(O4,O9,O14,O20)", changes)
    set_formula(ws, f"P{TOTAL_ROW}", "=SUM(P4,P9,P14,P20)", changes)


def hypercare_week_formula(row: int, *, sheet: str, col: str, port_ref: str, week_ref: str, day_ref: str | None = None) -> str:
    parts = [
        crit("'ROE_wk'!$AV:$AV", q("Ok")),
        crit("'ROE_wk'!$AY:$AY", port_ref),
        crit("'ROE_wk'!$AR:$AR", week_ref),
        crit("'ROE_wk'!$AI:$AI", q(ALLIANCE)),
        crit("'ROE_wk'!$N:$N", q("*HYPERCARE*")),
    ]
    if day_ref is not None:
        parts.insert(3, crit("'ROE_wk'!$AP:$AP", day_ref))
    return "=" + countifs(*parts)


def update_weekly_hypercare(ws, row: int, changes: list[dict]) -> None:
    for col in "FGHIJKL":
        set_formula(ws, f"{col}{row}", hypercare_week_formula(row, sheet=ws.Name, col=col, port_ref="$M$3", week_ref="$M$2", day_ref=f"{col}$8"), changes)
    set_formula(ws, f"M{row}", f"=SUM(F{row}:L{row})", changes)


def update_top_offenders(ws, changes: list[dict]) -> None:
    for row in range(20, 31):
        set_formula(ws, f"G{row}", f'=IFERROR(-(C{row}/B{row}-1),"")', changes)


def main() -> int:
    if not WORKBOOK.exists():
        print(f"Missing workbook: {WORKBOOK}")
        return 1
    changes: list[dict] = []
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    workbook = None
    try:
        workbook = call_with_retry(excel.Workbooks.Open, str(WORKBOOK), UpdateLinks=False, ReadOnly=False)
        call_with_retry(setattr, excel, "Calculation", -4135)  # manual

        for sheet in ("Month2date", "LastMonth_Overview"):
            update_month_sheet(workbook.Worksheets(sheet), changes)

        update_weekly_hypercare(workbook.Worksheets("Volume_DS"), 19, changes)
        update_weekly_hypercare(workbook.Worksheets("Volume_Graph"), 20, changes)
        update_weekly_hypercare(workbook.Worksheets("Volume_MAO"), 22, changes)
        update_top_offenders(workbook.Worksheets("Top_Offenders_Customers"), changes)

        call_with_retry(setattr, excel, "Calculation", -4105)  # automatic
        call_with_retry(excel.CalculateFullRebuild)
        call_with_retry(workbook.Save)
        call_with_retry(workbook.Close, SaveChanges=False)
        workbook = None
        REPORT.write_text(json.dumps({"workbook": str(WORKBOOK), "changes": changes}, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"Updated workbook: {WORKBOOK}")
        print(f"Changes: {len(changes)}")
        print(f"Report: {REPORT}")
        return 0
    except Exception as exc:  # noqa: BLE001
        print(f"Failed: {exc}")
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        return 1
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
