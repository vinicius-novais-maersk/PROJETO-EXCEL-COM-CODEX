from __future__ import annotations

import sys
from pathlib import Path

import pythoncom
import win32com.client

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from scripts.update_special_clients import WORKBOOK_PATH, backup_workbook, call_with_retry


PORT_ROWS = [2, 3, 6, 7, 8, 11, 12, 13, 16, 17, 18, 19, 20]
REGION_ROWS = [4, 9, 14, 21]
ALLIANCE = "Alian" + chr(231) + "a"


def log(message: str) -> None:
    print(f"[fix-oto-special] {message}", flush=True)


def q(value: str) -> str:
    return f'"{value}"'


def sheet_range(sheet_name: str, address: str) -> str:
    return f"'{sheet_name}'!{address}"


def criterion(range_ref: str, value: str) -> str:
    return f"{range_ref};{value}"


def countifs(*parts: str) -> str:
    return "CONT.SES(" + ";".join(parts) + ")"


def build_week_denominator(
    source_sheet: str,
    *,
    scope_column: str | None = None,
    scope_ref: str | None = None,
    alliance_filter: str | None = None,
) -> str:
    parts = [criterion(sheet_range(source_sheet, "$AV:$AV"), q("Ok"))]
    if scope_column and scope_ref:
        parts.append(criterion(sheet_range(source_sheet, f"${scope_column}:${scope_column}"), scope_ref))
    parts.append(criterion(sheet_range(source_sheet, "$AR:$AR"), "$AG$1"))
    if alliance_filter is not None:
        parts.append(criterion(sheet_range(source_sheet, "$AI:$AI"), q(alliance_filter)))
    parts.extend(
        [
            criterion(sheet_range(source_sheet, "$BB:$BB"), q("<>Sem Preenchimento")),
            criterion(sheet_range(source_sheet, "$BL:$BL"), q("N")),
            criterion(sheet_range(source_sheet, "$BO:$BO"), q("<>Especial")),
        ]
    )
    return countifs(*parts)


def build_week_late(
    source_sheet: str,
    *,
    scope_column: str | None = None,
    scope_ref: str | None = None,
    alliance_filter: str | None = None,
) -> str:
    parts = [criterion(sheet_range(source_sheet, "$AV:$AV"), q("Ok"))]
    if scope_column and scope_ref:
        parts.append(criterion(sheet_range(source_sheet, f"${scope_column}:${scope_column}"), scope_ref))
    parts.append(criterion(sheet_range(source_sheet, "$AR:$AR"), "$AG$1"))
    if alliance_filter is not None:
        parts.append(criterion(sheet_range(source_sheet, "$AI:$AI"), q(alliance_filter)))
    parts.extend(
        [
            criterion(sheet_range(source_sheet, "$BB:$BB"), q("Atrasado")),
            criterion(sheet_range(source_sheet, "$BL:$BL"), q("N")),
            criterion(sheet_range(source_sheet, "$BO:$BO"), q("<>Especial")),
        ]
    )
    return countifs(*parts)


def build_week_ratio_formula(
    source_sheet: str,
    *,
    scope_column: str | None = None,
    scope_ref: str | None = None,
    alliance_filter: str | None = None,
    rounded: bool = False,
) -> str:
    denominator = build_week_denominator(
        source_sheet,
        scope_column=scope_column,
        scope_ref=scope_ref,
        alliance_filter=alliance_filter,
    )
    late = build_week_late(
        source_sheet,
        scope_column=scope_column,
        scope_ref=scope_ref,
        alliance_filter=alliance_filter,
    )
    ratio = f'SE({denominator}=0;"";1-SEERRO({late}/{denominator};0))'
    if rounded:
        return f"=ARRED({ratio};2)"
    return f"={ratio}"


def build_port_lookup_formula(result_column: str) -> str:
    return (
        f'=SEERRO(PROCX($M$3;Week_Overview!AG:AG;Week_Overview!{result_column}:{result_column};"");"")'
    )


def build_historical_lookup_formula(week_cell: str, result_column: str) -> str:
    return (
        f'=SEERRO(PROCX($M$3;'
        f'INDIRETO("Week_"&${week_cell}&"!AG:AG");'
        f'INDIRETO("Week_"&${week_cell}&"!{result_column}:{result_column}");'
        f'"");"")'
    )


def build_port_denominator(
    alliance_filter: str | None,
    *,
    day_ref: str | None = None,
) -> str:
    parts = [
        criterion(sheet_range("ROE_wk", "$AV:$AV"), q("Ok")),
        criterion(sheet_range("ROE_wk", "$AY:$AY"), "$M$3"),
        criterion(sheet_range("ROE_wk", "$AR:$AR"), "$M$2"),
    ]
    if day_ref is not None:
        parts.append(criterion(sheet_range("ROE_wk", "$AP:$AP"), day_ref))
    if alliance_filter is not None:
        parts.append(criterion(sheet_range("ROE_wk", "$AI:$AI"), q(alliance_filter)))
    parts.extend(
        [
            criterion(sheet_range("ROE_wk", "$BB:$BB"), q("<>Sem Preenchimento")),
            criterion(sheet_range("ROE_wk", "$BL:$BL"), q("N")),
            criterion(sheet_range("ROE_wk", "$BO:$BO"), q("<>Especial")),
        ]
    )
    return countifs(*parts)


def build_port_late(
    alliance_filter: str | None,
    *,
    day_ref: str | None = None,
) -> str:
    parts = [
        criterion(sheet_range("ROE_wk", "$AV:$AV"), q("Ok")),
        criterion(sheet_range("ROE_wk", "$AY:$AY"), "$M$3"),
        criterion(sheet_range("ROE_wk", "$AR:$AR"), "$M$2"),
    ]
    if day_ref is not None:
        parts.append(criterion(sheet_range("ROE_wk", "$AP:$AP"), day_ref))
    if alliance_filter is not None:
        parts.append(criterion(sheet_range("ROE_wk", "$AI:$AI"), q(alliance_filter)))
    parts.extend(
        [
            criterion(sheet_range("ROE_wk", "$BB:$BB"), q("Atrasado")),
            criterion(sheet_range("ROE_wk", "$BL:$BL"), q("N")),
            criterion(sheet_range("ROE_wk", "$BO:$BO"), q("<>Especial")),
        ]
    )
    return countifs(*parts)


def build_port_ratio_formula(
    alliance_filter: str | None,
    *,
    day_ref: str | None = None,
    rounded: bool = True,
) -> str:
    denominator = build_port_denominator(alliance_filter, day_ref=day_ref)
    late = build_port_late(alliance_filter, day_ref=day_ref)
    ratio = f'SE({denominator}=0;"";1-SEERRO({late}/{denominator};0))'
    if rounded:
        return f'=SEERRO(ARRED({ratio};2);"")'
    return f"={ratio}"


def set_formula(worksheet, cell_ref: str, formula: str) -> None:
    call_with_retry(setattr, worksheet.Range(cell_ref), "FormulaLocal", formula)


def update_roe_late_flag(worksheet) -> None:
    data_body = call_with_retry(
        lambda: worksheet.ListObjects("ROE_wk").ListColumns("Atrasado?").DataBodyRange
    )
    formula = '=SE(E([@[OTD ajustado]]="Atrasado";[@[OTO Out]]="N";[@Especiais]<>"Especial");1;0)'
    call_with_retry(setattr, data_body, "FormulaLocal", formula)


def update_week_sheet(worksheet, source_sheet: str) -> int:
    changed = 0
    for row in PORT_ROWS:
        set_formula(
            worksheet,
            f"N{row}",
            build_week_ratio_formula(
                source_sheet,
                scope_column="AY",
                scope_ref=f"$AG{row}",
                alliance_filter=ALLIANCE,
            ),
        )
        set_formula(
            worksheet,
            f"P{row}",
            build_week_ratio_formula(
                source_sheet,
                scope_column="AY",
                scope_ref=f"$AG{row}",
                alliance_filter=f"<>{ALLIANCE}",
            ),
        )
        set_formula(
            worksheet,
            f"Q{row}",
            build_week_ratio_formula(
                source_sheet,
                scope_column="AY",
                scope_ref=f"$AG{row}",
            ),
        )
        changed += 3

    for row in REGION_ROWS:
        set_formula(
            worksheet,
            f"N{row}",
            build_week_ratio_formula(
                source_sheet,
                scope_column="AZ",
                scope_ref=f"$A{row}",
                alliance_filter=ALLIANCE,
            ),
        )
        set_formula(
            worksheet,
            f"P{row}",
            build_week_ratio_formula(
                source_sheet,
                scope_column="AZ",
                scope_ref=f"$A{row}",
                alliance_filter=f"<>{ALLIANCE}",
            ),
        )
        set_formula(
            worksheet,
            f"Q{row}",
            build_week_ratio_formula(
                source_sheet,
                scope_column="AZ",
                scope_ref=f"$A{row}",
            ),
        )
        changed += 3

    set_formula(
        worksheet,
        "N23",
        build_week_ratio_formula(source_sheet, alliance_filter=ALLIANCE, rounded=True),
    )
    set_formula(
        worksheet,
        "P23",
        build_week_ratio_formula(source_sheet, alliance_filter=f"<>{ALLIANCE}", rounded=True),
    )
    set_formula(
        worksheet,
        "Q23",
        build_week_ratio_formula(source_sheet),
    )
    return changed + 3


def update_volume_ds(worksheet) -> int:
    changed = 0
    for day_col in ["F", "G", "H", "I", "J", "K", "L"]:
        set_formula(
            worksheet,
            f"{day_col}16",
            build_port_ratio_formula(f"<>{ALLIANCE}", day_ref=f"{day_col}$8"),
        )
        changed += 1

    set_formula(worksheet, "M16", build_port_ratio_formula(f"<>{ALLIANCE}"))
    changed += 1

    for row in [25, 26]:
        set_formula(worksheet, f"J{row}", build_historical_lookup_formula(f"E{row}", "N"))
        set_formula(worksheet, f"N{row}", build_historical_lookup_formula(f"E{row}", "P"))
        set_formula(worksheet, f"R{row}", build_historical_lookup_formula(f"E{row}", "Q"))
        changed += 3

    set_formula(worksheet, "J27", build_port_lookup_formula("N"))
    set_formula(worksheet, "N27", build_port_lookup_formula("P"))
    set_formula(worksheet, "R27", build_port_lookup_formula("Q"))
    changed += 3
    return changed


def update_volume_graph(worksheet) -> int:
    changed = 0
    for day_col in ["F", "G", "H", "I", "J", "K", "L"]:
        set_formula(
            worksheet,
            f"{day_col}17",
            build_port_ratio_formula(ALLIANCE, day_ref=f"{day_col}$8"),
        )
        changed += 1

    set_formula(worksheet, "M17", build_port_ratio_formula(ALLIANCE))
    changed += 1

    for row in [25, 26]:
        set_formula(worksheet, f"J{row}", build_historical_lookup_formula(f"E{row}", "N"))
        set_formula(worksheet, f"N{row}", build_historical_lookup_formula(f"E{row}", "P"))
        set_formula(worksheet, f"R{row}", build_historical_lookup_formula(f"E{row}", "Q"))
        changed += 3

    set_formula(worksheet, "J27", build_port_lookup_formula("N"))
    set_formula(worksheet, "N27", build_port_lookup_formula("P"))
    set_formula(worksheet, "R27", build_port_lookup_formula("Q"))
    changed += 3
    return changed


def update_volume_mao(worksheet) -> int:
    changed = 0
    for day_col in ["F", "G", "H", "I", "J", "K", "L"]:
        set_formula(
            worksheet,
            f"{day_col}19",
            build_port_ratio_formula(ALLIANCE, day_ref=f"{day_col}$8"),
        )
        changed += 1

    set_formula(worksheet, "M19", build_port_ratio_formula(ALLIANCE))
    changed += 1

    for row in [27, 28]:
        set_formula(worksheet, f"J{row}", build_historical_lookup_formula(f"E{row}", "N"))
        set_formula(worksheet, f"N{row}", build_historical_lookup_formula(f"E{row}", "P"))
        set_formula(worksheet, f"R{row}", build_historical_lookup_formula(f"E{row}", "Q"))
        changed += 3

    set_formula(worksheet, "J29", build_port_lookup_formula("N"))
    set_formula(worksheet, "N29", build_port_lookup_formula("P"))
    set_formula(worksheet, "R29", build_port_lookup_formula("Q"))
    changed += 3
    return changed


def try_set_pivot_filter(pivot_table, field_name: str, target_value: str, *, allow_add: bool) -> bool:
    try:
        field = pivot_table.PivotFields(field_name)
    except Exception:
        return False
    if field.Orientation != 3:
        if not allow_add:
            return False
        try:
            field.Orientation = 3
            field.Position = 1
        except Exception:
            return False
    try:
        field.ClearAllFilters()
    except Exception:
        pass
    try:
        field.EnableMultiplePageItems = False
    except Exception:
        pass
    try:
        field.CurrentPage = target_value
        return True
    except Exception:
        return False


def normalize_pivot_page_value(value: int | float | str | None) -> str | None:
    if value in (None, ""):
        return None
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def refresh_top_offenders(worksheet, active_week: int | str | None) -> None:
    pivot_tables = worksheet.PivotTables()
    week_filter = normalize_pivot_page_value(active_week)
    for index in range(1, pivot_tables.Count + 1):
        pivot_table = pivot_tables.Item(index)
        try_set_pivot_filter(pivot_table, "Atrasado?", "1", allow_add=True)
        try_set_pivot_filter(pivot_table, "OTO Out", "N", allow_add=True)
        if week_filter is not None:
            try_set_pivot_filter(pivot_table, "Weeknum", week_filter, allow_add=False)
        call_with_retry(pivot_table.RefreshTable)


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

        changed = {}
        update_roe_late_flag(workbook.Worksheets("ROE_wk"))
        changed["ROE_wk[Atrasado?]"] = 1

        changed["Week_Overview"] = update_week_sheet(workbook.Worksheets("Week_Overview"), "ROE_wk")

        changed["Volume_DS"] = update_volume_ds(workbook.Worksheets("Volume_DS"))
        changed["Volume_Graph"] = update_volume_graph(workbook.Worksheets("Volume_Graph"))
        changed["Volume_MAO"] = update_volume_mao(workbook.Worksheets("Volume_MAO"))

        active_week = call_with_retry(lambda: workbook.Worksheets("Week_Overview").Range("AG1").Value)
        refresh_top_offenders(workbook.Worksheets("Top_Offenders_Customers"), active_week)
        refresh_top_offenders(workbook.Worksheets("Top_Offenders_Vendors"), active_week)

        call_with_retry(setattr, excel, "Calculation", -4105)
        call_with_retry(excel.CalculateFullRebuild)
        call_with_retry(workbook.Save)

        checks = {
            "Week_Overview!Q4": call_with_retry(lambda: workbook.Worksheets("Week_Overview").Range("Q4").Value),
            "Week_Overview!Q23": call_with_retry(lambda: workbook.Worksheets("Week_Overview").Range("Q23").Value),
            "Volume_DS!M16": call_with_retry(lambda: workbook.Worksheets("Volume_DS").Range("M16").Value),
            "Volume_Graph!M17": call_with_retry(lambda: workbook.Worksheets("Volume_Graph").Range("M17").Value),
            "Volume_MAO!M19": call_with_retry(lambda: workbook.Worksheets("Volume_MAO").Range("M19").Value),
            "Volume_DS!R27": call_with_retry(lambda: workbook.Worksheets("Volume_DS").Range("R27").Value),
            "Volume_Graph!R27": call_with_retry(lambda: workbook.Worksheets("Volume_Graph").Range("R27").Value),
            "Volume_MAO!R29": call_with_retry(lambda: workbook.Worksheets("Volume_MAO").Range("R29").Value),
            "Week_Overview!Q4_formula": call_with_retry(
                lambda: workbook.Worksheets("Week_Overview").Range("Q4").FormulaLocal
            ),
            "Volume_Graph!M17_formula": call_with_retry(
                lambda: workbook.Worksheets("Volume_Graph").Range("M17").FormulaLocal
            ),
            "Volume_MAO!R29_formula": call_with_retry(
                lambda: workbook.Worksheets("Volume_MAO").Range("R29").FormulaLocal
            ),
        }

        log(f"Changes applied: {changed}")
        log(f"Validation: {checks}")
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
