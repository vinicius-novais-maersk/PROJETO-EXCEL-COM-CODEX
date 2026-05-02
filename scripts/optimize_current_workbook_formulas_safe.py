from __future__ import annotations

import sys
from pathlib import Path

import pythoncom
import win32com.client

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from scripts.update_special_clients import WORKBOOK_PATH, backup_workbook, call_with_retry


BI_FORMULA = (
    '=IF([@[Tipo Serviço]]="Transporte Rodoviário",'
    'IF(COUNTIFS(ROE_wk[Provedor],[@Provedor],ROE_wk[Tipo Serviço],[@[Tipo Serviço]],'
    'ROE_wk[Porto],[@Porto],ROE_wk[Weeknum],[@Weeknum],ROE_wk[Volume],"Ok",'
    'ROE_wk[OTO Out],"N",ROE_wk[OTD ajustado],"<>Sem Preenchimento",ROE_wk[Especiais],"<>Especial")=0,"",'
    '1-COUNTIFS(ROE_wk[OTD ajustado],"Atrasado",ROE_wk[Provedor],[@Provedor],'
    'ROE_wk[Tipo Serviço],[@[Tipo Serviço]],ROE_wk[Porto],[@Porto],ROE_wk[Weeknum],[@Weeknum],'
    'ROE_wk[Volume],"Ok",ROE_wk[OTO Out],"N",ROE_wk[Especiais],"<>Especial")/'
    'COUNTIFS(ROE_wk[Provedor],[@Provedor],ROE_wk[Tipo Serviço],[@[Tipo Serviço]],'
    'ROE_wk[Porto],[@Porto],ROE_wk[Weeknum],[@Weeknum],ROE_wk[Volume],"Ok",'
    'ROE_wk[OTO Out],"N",ROE_wk[OTD ajustado],"<>Sem Preenchimento",ROE_wk[Especiais],"<>Especial")),"")'
)

BJ_FORMULA = (
    '=IF([@[Tipo Serviço]]="Transporte Rodoviário",'
    'IF(COUNTIFS(ROE_wk[Cliente Proposta],[@[Cliente Proposta]],ROE_wk[Tipo Serviço],[@[Tipo Serviço]],'
    'ROE_wk[Porto],[@Porto],ROE_wk[Weeknum],[@Weeknum],ROE_wk[Volume],"Ok",'
    'ROE_wk[OTO Out],"N",ROE_wk[OTD ajustado],"<>Sem Preenchimento",ROE_wk[Especiais],"<>Especial")=0,"",'
    '1-COUNTIFS(ROE_wk[OTD ajustado],"Atrasado",ROE_wk[Cliente Proposta],[@[Cliente Proposta]],'
    'ROE_wk[Tipo Serviço],[@[Tipo Serviço]],ROE_wk[Porto],[@Porto],ROE_wk[Weeknum],[@Weeknum],'
    'ROE_wk[Volume],"Ok",ROE_wk[OTO Out],"N",ROE_wk[Especiais],"<>Especial")/'
    'COUNTIFS(ROE_wk[Cliente Proposta],[@[Cliente Proposta]],ROE_wk[Tipo Serviço],[@[Tipo Serviço]],'
    'ROE_wk[Porto],[@Porto],ROE_wk[Weeknum],[@Weeknum],ROE_wk[Volume],"Ok",'
    'ROE_wk[OTO Out],"N",ROE_wk[OTD ajustado],"<>Sem Preenchimento",ROE_wk[Especiais],"<>Especial")),"")'
)

BK_FORMULA = (
    '=IF([@[Tipo Serviço]]="Transporte Rodoviário",'
    'IF(COUNTIFS(ROE_wk[Cliente Proposta],[@[Cliente Proposta]],ROE_wk[Tipo Serviço],[@[Tipo Serviço]],'
    'ROE_wk[Porto],[@Porto],ROE_wk[Volume],"Ok")=0,"",'
    '1-COUNTIFS(ROE_wk[SLA Ag],"N",ROE_wk[Cliente Proposta],[@[Cliente Proposta]],'
    'ROE_wk[Tipo Serviço],[@[Tipo Serviço]],ROE_wk[Porto],[@Porto],ROE_wk[Volume],"Ok")/'
    'COUNTIFS(ROE_wk[Cliente Proposta],[@[Cliente Proposta]],ROE_wk[Tipo Serviço],[@[Tipo Serviço]],'
    'ROE_wk[Porto],[@Porto],ROE_wk[Volume],"Ok")),"")'
)

CF_FORMULA = (
    '=1-COUNTIFS(ROE_wk[Volume],"Ok",ROE_wk[SLA Ag],"N",ROE_wk[Porto],"<>Manaus",'
    'ROE_wk[Centro de Custo],"Aliança")/'
    'COUNTIFS(ROE_wk[Volume],"Ok",ROE_wk[Porto],"<>Manaus",ROE_wk[Centro de Custo],"Aliança")'
)

MONTHLY_BJ_FORMULA = '=XLOOKUP(A2,ROE_wk[Nº OS],ROE_wk[% OTO Provider],"")'


def set_table_column_formula(list_object, column_name: str, formula: str) -> None:
    data_range = call_with_retry(lambda: list_object.ListColumns(column_name).DataBodyRange)
    call_with_retry(setattr, data_range, "Formula", formula)


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

        roe_ws = workbook.Worksheets("ROE_wk")
        roe_table = call_with_retry(lambda: roe_ws.ListObjects("ROE_wk"))
        set_table_column_formula(roe_table, "% OTO Provider", BI_FORMULA)
        set_table_column_formula(roe_table, "% OTO Client", BJ_FORMULA)
        set_table_column_formula(roe_table, "% SLA agend", BK_FORMULA)
        set_table_column_formula(roe_table, "% agend on time", CF_FORMULA)

        monthly_ws = workbook.Worksheets("ROE_wk_monthly")
        last_row = call_with_retry(lambda: monthly_ws.Cells(monthly_ws.Rows.Count, 1).End(-4162).Row)
        if last_row >= 2:
            monthly_range = monthly_ws.Range(f"BJ2:BJ{last_row}")
            call_with_retry(setattr, monthly_range, "Formula", MONTHLY_BJ_FORMULA)

        call_with_retry(setattr, excel, "Calculation", -4105)
        call_with_retry(excel.CalculateFullRebuild)
        call_with_retry(workbook.Save)
        call_with_retry(workbook.Close, SaveChanges=False)
        workbook = None

        print(f"Optimized formulas in ROE_wk and ROE_wk_monthly through row {last_row}.")
        return 0
    except Exception as exc:  # noqa: BLE001
        print(f"Failed to optimize formulas: {exc}")
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
    raise SystemExit(main())
