from __future__ import annotations

import sys
from pathlib import Path

import pythoncom
import win32com.client

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from scripts.update_special_clients import WORKBOOK_PATH, backup_workbook, call_with_retry


PORT_ROWS = (2, 3, 6, 7, 8, 11, 12, 13, 16, 17, 18, 19, 20)
REGION_ROWS = (4, 9, 14, 21)
ALL_R_ROWS = PORT_ROWS + REGION_ROWS + (23,)


def set_formula(ws, cell_ref: str, formula: str) -> None:
    call_with_retry(setattr, ws.Range(cell_ref), "Formula", formula)


def set_value(ws, cell_ref: str, value) -> None:
    call_with_retry(setattr, ws.Range(cell_ref), "Value", value)


def week_overview_fixes(ws) -> None:
    for row in PORT_ROWS:
        set_formula(
            ws,
            f"S{row}",
            (
                f'=SUMIFS(\'Reagendas\'!Q:Q,\'Reagendas\'!$F:$F,Week_Overview!AG{row},'
                f'\'Reagendas\'!$R:$R,$AI$1,\'Reagendas\'!$K:$K,"Reagenda")'
            ),
        )

        set_formula(
            ws,
            f"W{row}",
            (
                '=COUNTIFS(Validacao_Final!O:O,"<>Não",'
                'Validacao_Final!A:A,"Landside",'
                'Validacao_Final!J:J,"Reagendamento",'
                f'Validacao_Final!I:I,$AG{row},'
                'Validacao_Final!M:M,$AG$1)'
            ),
        )

        set_formula(
            ws,
            f"X{row}",
            (
                '=COUNTIFS(Validacao_Final!O:O,"<>Não",'
                'Validacao_Final!A:A,"Landside",'
                'Validacao_Final!J:J,"No-Show",'
                f'Validacao_Final!I:I,$AG{row},'
                'Validacao_Final!M:M,$AG$1)'
                '+COUNTIFS(Reagendas_2!BD:BD,"Negativo",'
                'Reagendas_2!AX:AX,"Landside",'
                f'Reagendas_2!AV:AV,$AG{row},'
                'Reagendas_2!AY:AY,$AG$1)'
            ),
        )

        set_formula(
            ws,
            f"Y{row}",
            (
                '=COUNTIFS(Reagendas_2!AX:AX,"CX",'
                f'Reagendas_2!AV:AV,$AG{row},'
                'Reagendas_2!AY:AY,$AG$1,'
                'Reagendas_2!AP:AP,"Agendamento - Falha de CX - Aliança")'
                '+COUNTIFS(Reagendas_2!AX:AX,"CX",'
                f'Reagendas_2!AV:AV,$AG{row},'
                'Reagendas_2!AY:AY,$AG$1,'
                'Reagendas_2!AP:AP,"Agendamento - Erro de Doc - Aliança")'
            ),
        )

        set_formula(
            ws,
            f"Z{row}",
            (
                '=COUNTIFS(Reagendas_2!AX:AX,"Outros",'
                f'Reagendas_2!AV:AV,$AG{row},'
                'Reagendas_2!AY:AY,$AG$1,'
                'Reagendas_2!AP:AP,"Agendamento - Falta de janela - Aliança")'
                '+COUNTIFS(Reagendas_2!AX:AX,"Outros",'
                f'Reagendas_2!AV:AV,$AG{row},'
                'Reagendas_2!AY:AY,$AG$1,'
                'Reagendas_2!AP:AP,"Erro Sistema – Aliança")'
                '+COUNTIFS(Reagendas_2!AX:AX,"Outros",'
                f'Reagendas_2!AV:AV,$AG{row},'
                'Reagendas_2!AY:AY,$AG$1,'
                'Reagendas_2!AP:AP,"Agendamento - Indisp de equip - Aliança")'
                '+COUNTIFS(Reagendas_2!AX:AX,"Outros",'
                f'Reagendas_2!AV:AV,$AG{row},'
                'Reagendas_2!AY:AY,$AG$1,'
                'Reagendas_2!AP:AP,"Agendamento - Atraso do navio - Aliança")'
                '+COUNTIFS(Reagendas_2!AX:AX,"Outros",'
                f'Reagendas_2!AV:AV,$AG{row},'
                'Reagendas_2!AY:AY,$AG$1,'
                'Reagendas_2!AP:AP,"Agendamento - Falha Portuária - Aliança")'
            ),
        )

        set_formula(
            ws,
            f"AA{row}",
            (
                '=COUNTIFS(Reagendas_2!AX:AX,"Outros",'
                f'Reagendas_2!AV:AV,$AG{row},'
                'Reagendas_2!AY:AY,$AG$1,'
                'Reagendas_2!AP:AP,"Agendamento - Indisp. Veículo - Aliança")'
            ),
        )

        set_formula(
            ws,
            f"AB{row}",
            (
                '=COUNTIFS(Reagendas_2!AX:AX,"Cliente",'
                f'Reagendas_2!AV:AV,$AG{row},'
                'Reagendas_2!AY:AY,$AG$1,'
                'Reagendas_2!AP:AP,"Agendamento - Erro de Doc - Cliente")'
                '+COUNTIFS(Reagendas_2!AX:AX,"Cliente",'
                f'Reagendas_2!AV:AV,$AG{row},'
                'Reagendas_2!AY:AY,$AG$1,'
                'Reagendas_2!AP:AP,"Agendamento - Reagendamento - Cliente")'
            ),
        )

        set_formula(
            ws,
            f"AC{row}",
            (
                '=COUNTIFS(Reagendas_2!AX:AX,"Landside",'
                f'Reagendas_2!AV:AV,$AG{row},'
                'Reagendas_2!AY:AY,$AG$1,'
                'Reagendas_2!BC:BC,"Negativo")'
            ),
        )

    for row in ALL_R_ROWS:
        if row == 23:
            formula = (
                '=IF(D23=0,"",1-IFERROR(COUNTIFS('
                '\'ROE_wk\'!$AV:$AV,"Ok",'
                '\'ROE_wk\'!$AR:$AR,$AG$1,'
                '\'ROE_wk\'!$AI:$AI,"Aliança",'
                '\'ROE_wk\'!$AT:$AT,"N")/D23,0))'
            )
        elif row in REGION_ROWS:
            formula = (
                f'=IF(D{row}=0,"",1-IFERROR(COUNTIFS('
                '\'ROE_wk\'!$AV:$AV,"Ok",'
                f'\'ROE_wk\'!$AZ:$AZ,$A{row},'
                '\'ROE_wk\'!$AR:$AR,$AG$1,'
                '\'ROE_wk\'!$AI:$AI,"Aliança",'
                '\'ROE_wk\'!$AT:$AT,"N")/D'
                f'{row},0))'
            )
        else:
            formula = (
                f'=IF(D{row}=0,"",1-IFERROR(COUNTIFS('
                '\'ROE_wk\'!$AV:$AV,"Ok",'
                f'\'ROE_wk\'!$AY:$AY,$AG{row},'
                '\'ROE_wk\'!$AR:$AR,$AG$1,'
                '\'ROE_wk\'!$AI:$AI,"Aliança",'
                '\'ROE_wk\'!$AT:$AT,"N")/D'
                f'{row},0))'
            )
        set_formula(ws, f"R{row}", formula)


def volume_ds_fixes(ws) -> None:
    set_value(ws, "E18", "Rescheduled")
    for col in "FGHIJKL":
        set_formula(
            ws,
            f"{col}18",
            (
                f'=SUMIFS(\'Reagendas\'!$Q:$Q,\'Reagendas\'!$F:$F,$M$3,'
                f'\'Reagendas\'!$R:$R,$M$2,\'Reagendas\'!$O:$O,{col}$8,'
                '\'Reagendas\'!$K:$K,"Reagenda")'
            ),
        )
    set_formula(ws, "M18", "=SUM(F18:L18)")


def volume_graph_fixes(ws, row: int) -> None:
    set_value(ws, f"E{row}", "Rescheduled")
    for col in "FGHIJKL":
        set_formula(
            ws,
            f"{col}{row}",
            (
                '=COUNTIFS('
                'Validacao_Final!$A:$A,"Landside",'
                'Validacao_Final!$I:$I,$M$3,'
                'Validacao_Final!$M:$M,$M$2,'
                f'Validacao_Final!$Q:$Q,{col}$8,'
                'Validacao_Final!$J:$J,"Reagendamento",'
                'Validacao_Final!$O:$O,"<>Não")'
            ),
        )
    set_formula(ws, f"M{row}", f"=SUM(F{row}:L{row})")


def main() -> int:
    backup_path = backup_workbook(WORKBOOK_PATH)
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
            str(WORKBOOK_PATH),
            UpdateLinks=False,
            ReadOnly=False,
        )

        call_with_retry(setattr, excel, "Calculation", -4135)

        week_overview_fixes(workbook.Worksheets("Week_Overview"))
        volume_ds_fixes(workbook.Worksheets("Volume_DS"))
        volume_graph_fixes(workbook.Worksheets("Volume_Graph"), 19)
        volume_graph_fixes(workbook.Worksheets("Volume_MAO"), 21)

        call_with_retry(setattr, excel, "Calculation", -4105)
        call_with_retry(excel.CalculateFullRebuild)
        call_with_retry(workbook.Save)
        call_with_retry(workbook.Close, SaveChanges=False)
        workbook = None

        print("Secondary weekly metrics aligned to the active week and relabeled.")
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
