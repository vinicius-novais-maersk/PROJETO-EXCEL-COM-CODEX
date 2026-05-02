from __future__ import annotations

import sys
from pathlib import Path

import win32com.client

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from scripts.update_special_clients import (
    WORKBOOK_PATH,
    backup_workbook,
    call_with_retry,
)

PORT_ROWS = [2, 3, 6, 7, 8, 11, 12, 13, 16, 17, 18, 19, 20]
REGION_ROWS = [4, 9, 14, 21]


def log(message: str) -> None:
    print(f"[exception-rule] {message}", flush=True)


def set_formula(worksheet, cell_ref: str, formula: str) -> None:
    call_with_retry(setattr, worksheet.Range(cell_ref), "FormulaLocal", formula)


def build_waiting_list_formula(result_column: str) -> str:
    return (
        f"=SEERRO(ÍNDICE(FILTRO('ROE_wk'!{result_column}:{result_column};"
        "(('ROE_wk'!BM:BM<>\"\")*('ROE_wk'!AY:AY=M3)));SEQUÊNCIA(12);1);\"\")"
    )


def build_week_overview_port_volume_formula(row: int, alliance_filter: str) -> str:
    return (
        "=SEERRO(CONT.SES("
        "'ROE_wk'!$AV:$AV;\"Ok\";"
        f"'ROE_wk'!$AY:$AY;$AG{row};"
        "'ROE_wk'!$AR:$AR;$AG$1;"
        f"'ROE_wk'!$AI:$AI;\"{alliance_filter}\""
        ');"")'
    )


def build_week_overview_port_total_formula(row: int) -> str:
    return (
        "=CONT.SES("
        "'ROE_wk'!$AV:$AV;\"Ok\";"
        f"'ROE_wk'!$AY:$AY;$AG{row};"
        "'ROE_wk'!$AR:$AR;$AG$1"
        ")"
    )


def build_week_overview_port_oto_formula(
    row: int, base_cell: str, alliance_filter: str
) -> str:
    return (
        f'=SE({base_cell}=0;"";1-SEERRO(CONT.SES('
        "'ROE_wk'!$AV:$AV;\"Ok\";"
        f"'ROE_wk'!$AY:$AY;$AG{row};"
        "'ROE_wk'!$AR:$AR;$AG$1;"
        f"'ROE_wk'!$AI:$AI;\"{alliance_filter}\";"
        "'ROE_wk'!$BB:$BB;\"Atrasado\";"
        "'ROE_wk'!$BL:$BL;\"N\""
        f")/{base_cell};0))"
    )


def build_week_overview_port_total_oto_formula(row: int, base_cell: str) -> str:
    return (
        f'=SE({base_cell}=0;"";1-SEERRO(CONT.SES('
        "'ROE_wk'!$AV:$AV;\"Ok\";"
        f"'ROE_wk'!$AY:$AY;$AG{row};"
        "'ROE_wk'!$AR:$AR;$AG$1;"
        "'ROE_wk'!$BB:$BB;\"Atrasado\";"
        "'ROE_wk'!$BL:$BL;\"N\""
        f")/{base_cell};0))"
    )


def build_week_overview_region_oto_formula(
    row: int, base_cell: str, alliance_filter: str
) -> str:
    return (
        f'=SE({base_cell}=0;"";1-SEERRO(CONT.SES('
        "'ROE_wk'!$AV:$AV;\"Ok\";"
        f"'ROE_wk'!$AZ:$AZ;$A{row};"
        "'ROE_wk'!$AR:$AR;$AG$1;"
        f"'ROE_wk'!$AI:$AI;\"{alliance_filter}\";"
        "'ROE_wk'!$BB:$BB;\"Atrasado\";"
        "'ROE_wk'!$BL:$BL;\"N\""
        f")/{base_cell};0))"
    )


def build_week_overview_region_total_oto_formula(row: int, base_cell: str) -> str:
    return (
        f'=SE({base_cell}=0;"";1-SEERRO(CONT.SES('
        "'ROE_wk'!$AV:$AV;\"Ok\";"
        f"'ROE_wk'!$AZ:$AZ;$A{row};"
        "'ROE_wk'!$AR:$AR;$AG$1;"
        "'ROE_wk'!$BB:$BB;\"Atrasado\";"
        "'ROE_wk'!$BL:$BL;\"N\""
        f")/{base_cell};0))"
    )


def build_week_overview_total_oto_formula(
    base_cell: str, alliance_filter: str, rounded: bool
) -> str:
    formula = (
        f'SE({base_cell}=0;"";1-SEERRO(CONT.SES('
        "'ROE_wk'!$AV:$AV;\"Ok\";"
        "'ROE_wk'!$AR:$AR;$AG$1;"
        f"'ROE_wk'!$AI:$AI;\"{alliance_filter}\";"
        "'ROE_wk'!$BB:$BB;\"Atrasado\";"
        "'ROE_wk'!$BL:$BL;\"N\""
        f")/{base_cell};0))"
    )
    if rounded:
        return f"=ARRED({formula};2)"
    return f"={formula}"


def build_week_overview_total_all_oto_formula(base_cell: str) -> str:
    return (
        f'=SE({base_cell}=0;"";1-SEERRO(CONT.SES('
        "'ROE_wk'!$AV:$AV;\"Ok\";"
        "'ROE_wk'!$AR:$AR;$AG$1;"
        "'ROE_wk'!$BB:$BB;\"Atrasado\";"
        "'ROE_wk'!$BL:$BL;\"N\""
        f")/{base_cell};0))"
    )


def build_week_overview_u_formula(row: int) -> str:
    return (
        "=SOMA(CONT.SES("
        "'ROE_wk'!$AV:$AV;\"Ok\";"
        f"'ROE_wk'!$AY:$AY;$AG{row};"
        "'ROE_wk'!$AR:$AR;$AG$1;"
        "'ROE_wk'!$BB:$BB;\"Atrasado\";"
        "'ROE_wk'!$BL:$BL;\"N\";"
        "'ROE_wk'!$AI:$AI;\"Aliança\""
        f");W{row};X{row})"
    )


def build_week_overview_v_formula(row: int) -> str:
    return (
        "=SOMA(CONT.SES("
        "'ROE_wk'!$AV:$AV;\"Ok\";"
        f"'ROE_wk'!$AY:$AY;$AG{row};"
        "'ROE_wk'!$AR:$AR;$AG$1;"
        "'ROE_wk'!$AI:$AI;\"Aliança\""
        f");W{row};X{row})"
    )


def build_daily_volume_formula(day_col: str, alliance_filter: str | None = None) -> str:
    criteria = (
        "'ROE_wk'!$AV:$AV;\"Ok\";"
        "'ROE_wk'!$AY:$AY;$M$3;"
        "'ROE_wk'!$AR:$AR;$M$2;"
        f"'ROE_wk'!$AP:$AP;{day_col}$8"
    )
    if alliance_filter is not None:
        criteria += f";'ROE_wk'!$AI:$AI;\"{alliance_filter}\""
    return f"=CONT.SES({criteria})"


def build_other_formula(day_col: str) -> str:
    return (
        "=CONT.SES("
        "'ROE_wk'!$AV:$AV;\"Ok\";"
        "'ROE_wk'!$AY:$AY;$M$3;"
        "'ROE_wk'!$AR:$AR;$M$2;"
        f"'ROE_wk'!$AP:$AP;{day_col}$8;"
        "'ROE_wk'!$AI:$AI;\"Maersk-Outros\""
        ")+CONT.SES("
        "'ROE_wk'!$AV:$AV;\"Ok\";"
        "'ROE_wk'!$AY:$AY;$M$3;"
        "'ROE_wk'!$AR:$AR;$M$2;"
        f"'ROE_wk'!$AP:$AP;{day_col}$8;"
        "'ROE_wk'!$AI:$AI;\"Outros\""
        ")"
    )


def build_hypercare_formula(day_col: str, alliance_filter: str) -> str:
    return (
        "=CONT.SES("
        "'ROE_wk'!$AV:$AV;\"Ok\";"
        "'ROE_wk'!$AY:$AY;$M$3;"
        "'ROE_wk'!$AR:$AR;$M$2;"
        f"'ROE_wk'!$AP:$AP;{day_col}$8;"
        f"'ROE_wk'!$AI:$AI;\"{alliance_filter}\";"
        "'ROE_wk'!$N:$N;\"*Reagenda*\""
        ")"
    )


def build_daily_oto_formula(day_col: str, base_ref: str, alliance_filter: str) -> str:
    return (
        f'=SE({base_ref}=0;"";ARRED(1-CONT.SES('
        "'ROE_wk'!$AV:$AV;\"Ok\";"
        "'ROE_wk'!$AY:$AY;$M$3;"
        "'ROE_wk'!$AR:$AR;$M$2;"
        f"'ROE_wk'!$AP:$AP;{day_col}$8;"
        f"'ROE_wk'!$AI:$AI;\"{alliance_filter}\";"
        "'ROE_wk'!$BB:$BB;\"Atrasado\";"
        "'ROE_wk'!$BL:$BL;\"N\""
        f")/{base_ref};2))"
    )


def build_weekly_oto_formula(base_ref: str, alliance_filter: str) -> str:
    return (
        f'=SE({base_ref}=0;"";ARRED(1-CONT.SES('
        "'ROE_wk'!$AV:$AV;\"Ok\";"
        "'ROE_wk'!$AY:$AY;$M$3;"
        "'ROE_wk'!$AR:$AR;$M$2;"
        f"'ROE_wk'!$AI:$AI;\"{alliance_filter}\";"
        "'ROE_wk'!$BB:$BB;\"Atrasado\";"
        "'ROE_wk'!$BL:$BL;\"N\""
        f")/{base_ref};2))"
    )


def build_daily_oto_mao_formula(day_col: str) -> str:
    base_ref = f"({day_col}12+{day_col}13)"
    return (
        f'=SE({base_ref}=0;"";ARRED(1-CONT.SES('
        "'ROE_wk'!$AV:$AV;\"Ok\";"
        "'ROE_wk'!$AY:$AY;$M$3;"
        "'ROE_wk'!$AR:$AR;$M$2;"
        f"'ROE_wk'!$AP:$AP;{day_col}$8;"
        "'ROE_wk'!$AI:$AI;\"Aliança\";"
        "'ROE_wk'!$BB:$BB;\"Atrasado\";"
        "'ROE_wk'!$BL:$BL;\"N\""
        f")/{base_ref};2))"
    )


def build_weekly_oto_mao_formula() -> str:
    base_ref = "(M12+M13)"
    return (
        f'=SE({base_ref}=0;"";ARRED(1-CONT.SES('
        "'ROE_wk'!$AV:$AV;\"Ok\";"
        "'ROE_wk'!$AY:$AY;$M$3;"
        "'ROE_wk'!$AR:$AR;$M$2;"
        "'ROE_wk'!$AI:$AI;\"Aliança\";"
        "'ROE_wk'!$BB:$BB;\"Atrasado\";"
        "'ROE_wk'!$BL:$BL;\"N\""
        f")/{base_ref};2))"
    )


def update_week_overview(worksheet) -> None:
    for row in PORT_ROWS:
        set_formula(
            worksheet,
            f"D{row}",
            build_week_overview_port_volume_formula(row, "Aliança"),
        )
        set_formula(
            worksheet,
            f"H{row}",
            build_week_overview_port_volume_formula(row, "<>Aliança"),
        )
        set_formula(worksheet, f"K{row}", build_week_overview_port_total_formula(row))
        set_formula(
            worksheet,
            f"N{row}",
            build_week_overview_port_oto_formula(row, f"D{row}", "Aliança"),
        )
        set_formula(
            worksheet,
            f"P{row}",
            build_week_overview_port_oto_formula(row, f"H{row}", "<>Aliança"),
        )
        set_formula(
            worksheet,
            f"Q{row}",
            build_week_overview_port_total_oto_formula(row, f"K{row}"),
        )
        set_formula(worksheet, f"U{row}", build_week_overview_u_formula(row))
        set_formula(worksheet, f"V{row}", build_week_overview_v_formula(row))

    for row in REGION_ROWS:
        set_formula(
            worksheet,
            f"N{row}",
            build_week_overview_region_oto_formula(row, f"D{row}", "Aliança"),
        )
        set_formula(
            worksheet,
            f"P{row}",
            build_week_overview_region_oto_formula(row, f"H{row}", "<>Aliança"),
        )
        set_formula(
            worksheet,
            f"Q{row}",
            build_week_overview_region_total_oto_formula(row, f"K{row}"),
        )

    set_formula(
        worksheet,
        "N23",
        build_week_overview_total_oto_formula("D23", "Aliança", rounded=True),
    )
    set_formula(
        worksheet,
        "P23",
        build_week_overview_total_oto_formula("H23", "<>Aliança", rounded=True),
    )
    set_formula(worksheet, "Q23", build_week_overview_total_all_oto_formula("K23"))


def update_volume_ds(worksheet) -> None:
    set_formula(worksheet, "O5", build_waiting_list_formula("BG"))
    set_formula(worksheet, "P5", build_waiting_list_formula("A"))
    set_formula(worksheet, "Q5", build_waiting_list_formula("BN"))

    for day_col in ["F", "G", "H", "I", "J", "K", "L"]:
        set_formula(worksheet, f"{day_col}9", build_daily_volume_formula(day_col))
        set_formula(
            worksheet, f"{day_col}11", build_daily_volume_formula(day_col, "Aliança")
        )
        set_formula(
            worksheet, f"{day_col}12", build_daily_volume_formula(day_col, "<>Aliança")
        )
        set_formula(worksheet, f"{day_col}13", build_other_formula(day_col))
        set_formula(
            worksheet,
            f"{day_col}16",
            build_daily_oto_formula(day_col, f"{day_col}12", "<>Aliança"),
        )
        set_formula(
            worksheet, f"{day_col}19", build_hypercare_formula(day_col, "Maersk")
        )

    set_formula(worksheet, "M16", build_weekly_oto_formula("M12", "<>Aliança"))


def update_volume_graph(worksheet) -> None:
    set_formula(worksheet, "O5", build_waiting_list_formula("BG"))
    set_formula(worksheet, "P5", build_waiting_list_formula("A"))
    set_formula(worksheet, "Q5", build_waiting_list_formula("BN"))

    for day_col in ["F", "G", "H", "I", "J", "K", "L"]:
        set_formula(worksheet, f"{day_col}9", build_daily_volume_formula(day_col))
        set_formula(
            worksheet, f"{day_col}12", build_daily_volume_formula(day_col, "Aliança")
        )
        set_formula(
            worksheet, f"{day_col}13", build_daily_volume_formula(day_col, "<>Aliança")
        )
        set_formula(worksheet, f"{day_col}14", build_other_formula(day_col))
        set_formula(
            worksheet,
            f"{day_col}17",
            build_daily_oto_formula(day_col, f"{day_col}12", "Aliança"),
        )
        set_formula(
            worksheet, f"{day_col}20", build_hypercare_formula(day_col, "Aliança")
        )

    set_formula(worksheet, "M17", build_weekly_oto_formula("M12", "Aliança"))


def update_volume_mao(worksheet) -> None:
    set_formula(worksheet, "O5", build_waiting_list_formula("BG"))
    set_formula(worksheet, "P5", build_waiting_list_formula("A"))
    set_formula(worksheet, "Q5", build_waiting_list_formula("BN"))

    for day_col in ["F", "G", "H", "I", "J", "K", "L"]:
        set_formula(worksheet, f"{day_col}9", build_daily_volume_formula(day_col))
        set_formula(
            worksheet, f"{day_col}14", build_daily_volume_formula(day_col, "<>Aliança")
        )
        set_formula(worksheet, f"{day_col}15", build_other_formula(day_col))
        set_formula(worksheet, f"{day_col}19", build_daily_oto_mao_formula(day_col))
        set_formula(
            worksheet, f"{day_col}22", build_hypercare_formula(day_col, "Aliança")
        )

    set_formula(worksheet, "M19", build_weekly_oto_mao_formula())


def update_roe_late_flag(worksheet) -> None:
    data_body = call_with_retry(
        lambda: worksheet.ListObjects("ROE_wk").ListColumns("Atrasado?").DataBodyRange
    )
    formula = '=SE(E([@[OTD ajustado]]="Atrasado";[@[OTO Out]]="N";[@Especiais]<>"Especial");1;0)'
    call_with_retry(setattr, data_body, "FormulaLocal", formula)


def try_set_pivot_filter(
    pivot_table,
    field_name: str,
    target_value: str,
    *,
    allow_add: bool,
) -> bool:
    field = pivot_table.PivotFields(field_name)
    if field.Orientation != 3:
        if not allow_add:
            return False
        try:
            field.Orientation = 3
            field.Position = 1
        except Exception:  # noqa: BLE001
            return False
    try:
        field.ClearAllFilters()
    except Exception:  # noqa: BLE001
        pass
    try:
        field.EnableMultiplePageItems = False
    except Exception:  # noqa: BLE001
        pass
    field.CurrentPage = target_value
    return True


def update_top_offenders(worksheet) -> None:
    pivot_tables = worksheet.PivotTables()
    for index in range(1, pivot_tables.Count + 1):
        pivot_table = pivot_tables.Item(index)
        try_set_pivot_filter(pivot_table, "Atrasado?", "1", allow_add=False)
        try_set_pivot_filter(pivot_table, "OTO Out", "N", allow_add=True)
        call_with_retry(pivot_table.RefreshTable)


def audit_remaining_special_usage(workbook) -> list[str]:
    findings: list[str] = []
    target_sheets = ["Volume_DS", "Volume_Graph", "Volume_MAO", "Week_Overview"]
    for sheet_name in target_sheets:
        worksheet = workbook.Worksheets(sheet_name)
        used = worksheet.UsedRange
        rows = used.Rows.Count
        cols = used.Columns.Count
        for row in range(1, rows + 1):
            for col in range(1, cols + 1):
                formula = worksheet.Cells(row, col).FormulaLocal
                if not isinstance(formula, str):
                    continue
                if "BO:$BO" not in formula:
                    continue
                findings.append(f"{sheet_name}!{worksheet.Cells(row, col).Address}")
    return findings


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
    excel.ScreenUpdating = False

    workbook = None
    try:
        workbook = call_with_retry(
            excel.Workbooks.Open,
            str(WORKBOOK_PATH),
            UpdateLinks=False,
            ReadOnly=False,
        )

        update_roe_late_flag(workbook.Worksheets("ROE_wk"))
        update_week_overview(workbook.Worksheets("Week_Overview"))
        update_volume_ds(workbook.Worksheets("Volume_DS"))
        update_volume_graph(workbook.Worksheets("Volume_Graph"))
        update_volume_mao(workbook.Worksheets("Volume_MAO"))
        update_top_offenders(workbook.Worksheets("Top_Offenders_Customers"))
        update_top_offenders(workbook.Worksheets("Top_Offenders_Vendors"))

        call_with_retry(setattr, excel, "Calculation", -4105)
        call_with_retry(excel.CalculateFullRebuild)
        for sheet_name in ["Week_Overview", "Volume_DS", "Volume_Graph", "Volume_MAO"]:
            call_with_retry(workbook.Worksheets(sheet_name).Calculate)

        call_with_retry(workbook.Save)

        findings = audit_remaining_special_usage(workbook)
        log(f"Remaining BO references in audited sheets: {len(findings)}")
        for item in findings[:20]:
            log(f"Remaining -> {item}")

        log(
            f"Week_Overview!Q16 -> {call_with_retry(lambda: workbook.Worksheets('Week_Overview').Range('Q16').Value)}"
        )
        log(
            f"Week_Overview!D16 -> {call_with_retry(lambda: workbook.Worksheets('Week_Overview').Range('D16').Value)}"
        )
        log(
            f"Week_Overview!H16 -> {call_with_retry(lambda: workbook.Worksheets('Week_Overview').Range('H16').Value)}"
        )
        log(
            f"Week_Overview!K16 -> {call_with_retry(lambda: workbook.Worksheets('Week_Overview').Range('K16').Value)}"
        )

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
        try:
            excel.Quit()
        except Exception:  # noqa: BLE001
            pass


if __name__ == "__main__":
    sys.exit(main())
