from __future__ import annotations

import os
import sys
from pathlib import Path

import pythoncom
import win32com.client

try:
    from scripts.update_special_clients import BACKUP_DIR, call_with_retry
except ModuleNotFoundError:
    from update_special_clients import BACKUP_DIR, call_with_retry

try:
    from scripts.workbook_paths import resolve_workbook_path
except ModuleNotFoundError:
    from workbook_paths import resolve_workbook_path


ALLIANCE = "Alian" + chr(231) + "a"
NO_WORD = "N" + chr(227) + "o"
SUN = "Sun"

PORT_ROWS = (2, 3, 6, 7, 8, 11, 12, 13, 16, 17, 18, 19, 20)
REGION_ROWS = (4, 9, 14, 21)
ALL_R_ROWS = PORT_ROWS + REGION_ROWS + (23,)


def backup_workbook(workbook_path: Path) -> Path:
    from datetime import datetime
    import shutil

    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = BACKUP_DIR / f"{workbook_path.stem}_backup_{timestamp}{workbook_path.suffix}"
    shutil.copy2(workbook_path, backup_path)
    return backup_path


def set_formula(ws, cell_ref: str, formula: str) -> None:
    call_with_retry(setattr, ws.Range(cell_ref), "Formula", formula)


def set_formula_local(ws, cell_ref: str, formula: str) -> None:
    call_with_retry(setattr, ws.Range(cell_ref), "FormulaLocal", formula)


def set_value(ws, cell_ref: str, value) -> None:
    call_with_retry(setattr, ws.Range(cell_ref), "Value", value)


def q(value: str) -> str:
    return f'"{value}"'


def build_countifs(*parts: str) -> str:
    return "COUNTIFS(" + ",".join(parts) + ")"


def build_sumifs(sum_range: str, *parts: str) -> str:
    return "SUMIFS(" + ",".join([sum_range, *parts]) + ")"


def criterion(range_ref: str, crit: str) -> str:
    return f"{range_ref},{crit}"


def build_not_no_criterion(range_ref: str) -> str:
    return f'{range_ref},"<>"&"N"&CHAR(227)&"o"'


def build_port_oto_formula(
    *,
    alliance_filter: str | None,
    port_ref: str,
    week_ref: str,
    day_ref: str | None = None,
) -> str:
    denominator_parts = [
        criterion("'ROE_wk'!$AV:$AV", q("Ok")),
        criterion("'ROE_wk'!$AY:$AY", port_ref),
        criterion("'ROE_wk'!$AR:$AR", week_ref),
    ]
    if day_ref is not None:
        denominator_parts.append(criterion("'ROE_wk'!$AP:$AP", day_ref))
    if alliance_filter is not None:
        denominator_parts.append(criterion("'ROE_wk'!$AI:$AI", q(alliance_filter)))
    denominator_parts.extend(
        [
            criterion("'ROE_wk'!$BB:$BB", q("<>Sem Preenchimento")),
            criterion("'ROE_wk'!$BL:$BL", q("N")),
        ]
    )
    denominator = build_countifs(*denominator_parts)

    late_parts = denominator_parts.copy()
    late_parts[-2] = criterion("'ROE_wk'!$BB:$BB", q("Atrasado"))
    late_parts.append(criterion("'ROE_wk'!$BO:$BO", q("<>Especial")))
    late = build_countifs(*late_parts)

    return f'=IFERROR(ROUND(IF({denominator}=0,"",1-IFERROR({late}/{denominator},0)),2),"")'


def build_s_formula(row: int) -> str:
    return (
        "="
        + build_sumifs(
            "'Reagendas'!$Q:$Q",
            criterion("'Reagendas'!$F:$F", f"$AG{row}"),
            criterion("'Reagendas'!$R:$R", "$AG$1"),
            criterion("'Reagendas'!$K:$K", q("Reagenda")),
        )
    )


def build_w_formula(row: int) -> str:
    return (
        "="
        + build_countifs(
            build_not_no_criterion("Validacao_Final!$O:$O"),
            criterion("Validacao_Final!$A:$A", q("Landside")),
            criterion("Validacao_Final!$J:$J", q("Reagendamento")),
            criterion("Validacao_Final!$I:$I", f"$AG{row}"),
            criterion("Validacao_Final!$M:$M", "$AG$1"),
        )
    )


def build_x_formula(row: int) -> str:
    validacao = build_countifs(
        build_not_no_criterion("Validacao_Final!$O:$O"),
        criterion("Validacao_Final!$A:$A", q("Landside")),
        criterion("Validacao_Final!$J:$J", q("No-Show")),
        criterion("Validacao_Final!$I:$I", f"$AG{row}"),
        criterion("Validacao_Final!$M:$M", "$AG$1"),
    )
    reagendas = build_countifs(
        criterion("Reagendas_2!$BD:$BD", q("Negativo")),
        criterion("Reagendas_2!$AX:$AX", q("Landside")),
        criterion("Reagendas_2!$AV:$AV", f"$AG{row}"),
        criterion("Reagendas_2!$AY:$AY", "$AG$1"),
    )
    return f"={validacao}+{reagendas}"


def build_y_formula(row: int) -> str:
    part1 = build_countifs(
        criterion("Reagendas_2!$AX:$AX", q("CX")),
        criterion("Reagendas_2!$AV:$AV", f"$AG{row}"),
        criterion("Reagendas_2!$AY:$AY", "$AG$1"),
        criterion("Reagendas_2!$AP:$AP", q("*Falha de CX*")),
    )
    part2 = build_countifs(
        criterion("Reagendas_2!$AX:$AX", q("CX")),
        criterion("Reagendas_2!$AV:$AV", f"$AG{row}"),
        criterion("Reagendas_2!$AY:$AY", "$AG$1"),
        criterion("Reagendas_2!$AP:$AP", q("*Erro de Doc*")),
    )
    return f"={part1}+{part2}"


def build_z_formula(row: int) -> str:
    patterns = (
        "*Falta de Janela*",
        "*Erro Sistema*",
        "*Equip*",
        "*Atraso do Navio*",
        "*Falha Portu*",
    )
    parts = [
        build_countifs(
            criterion("Reagendas_2!$AX:$AX", q("Outros")),
            criterion("Reagendas_2!$AV:$AV", f"$AG{row}"),
            criterion("Reagendas_2!$AY:$AY", "$AG$1"),
            criterion("Reagendas_2!$AP:$AP", q(pattern)),
        )
        for pattern in patterns
    ]
    return "=" + "+".join(parts)


def build_aa_formula(row: int) -> str:
    return (
        "="
        + build_countifs(
            criterion("Reagendas_2!$AX:$AX", q("Outros")),
            criterion("Reagendas_2!$AV:$AV", f"$AG{row}"),
            criterion("Reagendas_2!$AY:$AY", "$AG$1"),
            criterion("Reagendas_2!$AP:$AP", q("*Veic*")),
        )
    )


def build_ab_formula(row: int) -> str:
    part1 = build_countifs(
        criterion("Reagendas_2!$AX:$AX", q("Cliente")),
        criterion("Reagendas_2!$AV:$AV", f"$AG{row}"),
        criterion("Reagendas_2!$AY:$AY", "$AG$1"),
        criterion("Reagendas_2!$AP:$AP", q("*Erro de Doc*")),
    )
    part2 = build_countifs(
        criterion("Reagendas_2!$AX:$AX", q("Cliente")),
        criterion("Reagendas_2!$AV:$AV", f"$AG{row}"),
        criterion("Reagendas_2!$AY:$AY", "$AG$1"),
        criterion("Reagendas_2!$AP:$AP", q("*Reagendamento*")),
    )
    return f"={part1}+{part2}"


def build_ac_formula(row: int, diff_text: str) -> str:
    return (
        "="
        + build_countifs(
            criterion("Reagendas_2!$AX:$AX", q("Landside")),
            criterion("Reagendas_2!$AV:$AV", f"$AG{row}"),
            criterion("Reagendas_2!$AY:$AY", "$AG$1"),
            criterion("Reagendas_2!$BC:$BC", q(diff_text)),
        )
    )


def build_ae_formula(row: int) -> str:
    return (
        "="
        + build_countifs(
            criterion("Reagendas_2!$AX:$AX", q("Landside")),
            criterion("Reagendas_2!$AV:$AV", f"$AG{row}"),
            criterion("Reagendas_2!$AY:$AY", "$AG$1"),
            criterion("Reagendas_2!$BC:$BC", q("*Indisp*Veic*")),
        )
    )


def build_af_formula(row: int) -> str:
    return (
        "="
        + build_countifs(
            criterion("Reagendas_2!$AV:$AV", f"$AG{row}"),
            criterion("Reagendas_2!$AY:$AY", "$AG$1"),
        )
    )


def build_r_formula(row: int, *, region_ref: bool = False, total_ref: bool = False) -> str:
    if total_ref:
        criteria = [
            criterion("'ROE_wk'!$AV:$AV", q("Ok")),
            criterion("'ROE_wk'!$AR:$AR", "$AG$1"),
            criterion("'ROE_wk'!$AI:$AI", q(ALLIANCE)),
            criterion("'ROE_wk'!$AT:$AT", q("N")),
        ]
        numerator = build_countifs(*criteria)
        return f'=IF(D23=0,"",1-IFERROR({numerator}/D23,0))'

    scope_ref = f"$A{row}" if region_ref else f"$AG{row}"
    scope_col = "$AZ:$AZ" if region_ref else "$AY:$AY"
    criteria = [
        criterion("'ROE_wk'!$AV:$AV", q("Ok")),
        criterion(f"'ROE_wk'!{scope_col}", scope_ref),
        criterion("'ROE_wk'!$AR:$AR", "$AG$1"),
        criterion("'ROE_wk'!$AI:$AI", q(ALLIANCE)),
        criterion("'ROE_wk'!$AT:$AT", q("N")),
    ]
    numerator = build_countifs(*criteria)
    return f'=IF(D{row}=0,"",1-IFERROR({numerator}/D{row},0))'


def update_aux_and_day_week(workbook) -> None:
    aux_ws = workbook.Worksheets("Aux")
    set_value(aux_ws, "A7", 1)
    set_value(aux_ws, "B7", SUN)

    roe_ws = workbook.Worksheets("ROE_wk")
    day_body = call_with_retry(lambda: roe_ws.ListObjects("ROE_wk").ListColumns("day week").DataBodyRange)
    call_with_retry(setattr, day_body, "FormulaLocal", '=SEERRO(PROCX(AQ2;Aux!$A$1:$A$7;Aux!$B$1:$B$7;"");"")')

    monthly_ws = workbook.Worksheets("ROE_wk_monthly")
    last_monthly_row = call_with_retry(lambda: monthly_ws.Cells(monthly_ws.Rows.Count, 1).End(-4162).Row)
    monthly_range = monthly_ws.Range(f"AP2:AP{last_monthly_row}")
    call_with_retry(setattr, monthly_range, "FormulaLocal", '=SEERRO(PROCX(AQ2;Aux!$A$1:$A$7;Aux!$B$1:$B$7;"");"")')


def update_week_overview(ws) -> None:
    for row in PORT_ROWS:
        set_formula(ws, f"S{row}", build_s_formula(row))
        set_formula(ws, f"W{row}", build_w_formula(row))
        set_formula(ws, f"X{row}", build_x_formula(row))
        set_formula(ws, f"Y{row}", build_y_formula(row))
        set_formula(ws, f"Z{row}", build_z_formula(row))
        set_formula(ws, f"AA{row}", build_aa_formula(row))
        set_formula(ws, f"AB{row}", build_ab_formula(row))
        set_formula(ws, f"AC{row}", build_ac_formula(row, "Negativo"))
        set_formula(ws, f"AD{row}", build_ac_formula(row, "Sem OS"))
        set_formula(ws, f"AE{row}", build_ae_formula(row))
        set_formula(ws, f"AF{row}", build_af_formula(row))

    for row in PORT_ROWS:
        set_formula(ws, f"R{row}", build_r_formula(row))
    for row in REGION_ROWS:
        set_formula(ws, f"R{row}", build_r_formula(row, region_ref=True))
    set_formula(ws, "R23", build_r_formula(23, total_ref=True))


def update_volume_ds(ws) -> None:
    set_value(ws, "E18", "Rescheduled")
    for col in "FGHIJKL":
        set_formula(
            ws,
            f"{col}16",
            build_port_oto_formula(
                alliance_filter=f"<>{ALLIANCE}",
                port_ref="$M$3",
                week_ref="$M$2",
                day_ref=f"{col}$8",
            ),
        )
        set_formula(
            ws,
            f"{col}18",
            (
                "="
                + build_sumifs(
                    "'Reagendas'!$Q:$Q",
                    criterion("'Reagendas'!$F:$F", "$M$3"),
                    criterion("'Reagendas'!$R:$R", "$M$2"),
                    criterion("'Reagendas'!$O:$O", f"{col}$8"),
                    criterion("'Reagendas'!$K:$K", q("Reagenda")),
                )
            ),
        )
    set_formula(
        ws,
        "M16",
        build_port_oto_formula(
            alliance_filter=f"<>{ALLIANCE}",
            port_ref="$M$3",
            week_ref="$M$2",
        ),
    )
    set_formula(ws, "M18", "=SUM(F18:L18)")


def update_volume_graph(ws) -> None:
    for col in "FGHIJKL":
        set_formula(
            ws,
            f"{col}17",
            build_port_oto_formula(
                alliance_filter=ALLIANCE,
                port_ref="$M$3",
                week_ref="$M$2",
                day_ref=f"{col}$8",
            ),
        )
    set_formula(
        ws,
        "M17",
        build_port_oto_formula(
            alliance_filter=ALLIANCE,
            port_ref="$M$3",
            week_ref="$M$2",
        ),
    )


def update_volume_mao(ws) -> None:
    for col in "FGHIJKL":
        set_formula(
            ws,
            f"{col}19",
            build_port_oto_formula(
                alliance_filter=ALLIANCE,
                port_ref="$M$3",
                week_ref="$M$2",
                day_ref=f"{col}$8",
            ),
        )
    set_formula(
        ws,
        "M19",
        build_port_oto_formula(
            alliance_filter=ALLIANCE,
            port_ref="$M$3",
            week_ref="$M$2",
        ),
    )


def main() -> int:
    workbook_path = Path(os.environ.get("DSU_WORKBOOK_PATH", str(resolve_workbook_path())))
    if not workbook_path.exists():
        print(f"Workbook not found: {workbook_path}")
        return 1

    backup_path = backup_workbook(workbook_path)
    print(f"Backup created at: {backup_path}")

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

        update_aux_and_day_week(workbook)
        update_week_overview(workbook.Worksheets("Week_Overview"))
        update_volume_ds(workbook.Worksheets("Volume_DS"))
        update_volume_graph(workbook.Worksheets("Volume_Graph"))
        update_volume_mao(workbook.Worksheets("Volume_MAO"))

        call_with_retry(setattr, excel, "Calculation", -4105)
        call_with_retry(excel.CalculateFullRebuild)
        call_with_retry(workbook.Save)
        call_with_retry(workbook.Close, SaveChanges=False)
        workbook = None

        print("Confirmed dashboard fixes applied successfully.")
        return 0
    except Exception as exc:  # noqa: BLE001
        print(f"Failed to update workbook: {exc}")
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
