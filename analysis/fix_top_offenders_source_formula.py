from __future__ import annotations

import json
import shutil
import sys
from datetime import datetime
from pathlib import Path

import pythoncom
import win32com.client

try:
    from scripts.update_special_clients import call_with_retry
except ModuleNotFoundError:
    sys.path.insert(0, str(Path.cwd()))
    from scripts.update_special_clients import call_with_retry

WORKBOOK = Path(r"C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX\analysis\Base_DSU2026_WK19_rules_hypercare_fix_20260508_173020.xlsm")
STAMP = datetime.now().strftime("%Y%m%d_%H%M%S")
BACKUP = WORKBOOK.with_name(WORKBOOK.stem + f"_before_top_offenders_fix_{STAMP}" + WORKBOOK.suffix)
REPORT = WORKBOOK.with_name(WORKBOOK.stem + f"_top_offenders_source_fix_{STAMP}.json")

CLIENT_FORMULA = (
    '=IF(ROE_wk[[#This Row],[Tipo Serviço]]="Transporte Rodoviário",'
    'IF(COUNTIFS(ROE_wk[Cliente Proposta],ROE_wk[[#This Row],[Cliente Proposta]],'
    'ROE_wk[Tipo Serviço],ROE_wk[[#This Row],[Tipo Serviço]],'
    'ROE_wk[Porto],ROE_wk[[#This Row],[Porto]],'
    'ROE_wk[Weeknum],ROE_wk[[#This Row],[Weeknum]],'
    'ROE_wk[Volume],"Ok",ROE_wk[OTO Out],"N",ROE_wk[OTD ajustado],"<>Sem Preenchimento")=0,'
    '"",'
    '1-COUNTIFS(ROE_wk[OTD ajustado],"Atrasado",'
    'ROE_wk[Cliente Proposta],ROE_wk[[#This Row],[Cliente Proposta]],'
    'ROE_wk[Tipo Serviço],ROE_wk[[#This Row],[Tipo Serviço]],'
    'ROE_wk[Porto],ROE_wk[[#This Row],[Porto]],'
    'ROE_wk[Weeknum],ROE_wk[[#This Row],[Weeknum]],'
    'ROE_wk[Volume],"Ok",ROE_wk[OTO Out],"N",ROE_wk[Especiais],"<>Especial")/'
    'COUNTIFS(ROE_wk[Cliente Proposta],ROE_wk[[#This Row],[Cliente Proposta]],'
    'ROE_wk[Tipo Serviço],ROE_wk[[#This Row],[Tipo Serviço]],'
    'ROE_wk[Porto],ROE_wk[[#This Row],[Porto]],'
    'ROE_wk[Weeknum],ROE_wk[[#This Row],[Weeknum]],'
    'ROE_wk[Volume],"Ok",ROE_wk[OTO Out],"N",ROE_wk[OTD ajustado],"<>Sem Preenchimento")),"")'
)

PROVIDER_FORMULA = (
    '=IF(ROE_wk[[#This Row],[Tipo Serviço]]="Transporte Rodoviário",'
    'IF(COUNTIFS(ROE_wk[Provedor],ROE_wk[[#This Row],[Provedor]],'
    'ROE_wk[Tipo Serviço],ROE_wk[[#This Row],[Tipo Serviço]],'
    'ROE_wk[Porto],ROE_wk[[#This Row],[Porto]],'
    'ROE_wk[Weeknum],ROE_wk[[#This Row],[Weeknum]],'
    'ROE_wk[Volume],"Ok",ROE_wk[OTO Out],"N",ROE_wk[OTD ajustado],"<>Sem Preenchimento")=0,'
    '"",'
    '1-COUNTIFS(ROE_wk[OTD ajustado],"Atrasado",'
    'ROE_wk[Provedor],ROE_wk[[#This Row],[Provedor]],'
    'ROE_wk[Tipo Serviço],ROE_wk[[#This Row],[Tipo Serviço]],'
    'ROE_wk[Porto],ROE_wk[[#This Row],[Porto]],'
    'ROE_wk[Weeknum],ROE_wk[[#This Row],[Weeknum]],'
    'ROE_wk[Volume],"Ok",ROE_wk[OTO Out],"N",ROE_wk[Especiais],"<>Especial")/'
    'COUNTIFS(ROE_wk[Provedor],ROE_wk[[#This Row],[Provedor]],'
    'ROE_wk[Tipo Serviço],ROE_wk[[#This Row],[Tipo Serviço]],'
    'ROE_wk[Porto],ROE_wk[[#This Row],[Porto]],'
    'ROE_wk[Weeknum],ROE_wk[[#This Row],[Weeknum]],'
    'ROE_wk[Volume],"Ok",ROE_wk[OTO Out],"N",ROE_wk[OTD ajustado],"<>Sem Preenchimento")),"")'
)


def main() -> int:
    if not WORKBOOK.exists():
        print(f"Workbook not found: {WORKBOOK}")
        return 1
    shutil.copy2(WORKBOOK, BACKUP)
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    result = {"workbook": str(WORKBOOK), "backup": str(BACKUP), "timestamp": datetime.now().isoformat(timespec="seconds"), "changes": []}
    wb = None
    try:
        wb = call_with_retry(excel.Workbooks.Open, str(WORKBOOK), UpdateLinks=False, ReadOnly=False)
        call_with_retry(setattr, excel, "Calculation", -4135)  # manual
        lo = wb.Worksheets("ROE_wk").ListObjects("ROE_wk")
        for col_name, formula in [("% OTO Client", CLIENT_FORMULA), ("% OTO Provider", PROVIDER_FORMULA)]:
            body = call_with_retry(lambda c=col_name: lo.ListColumns(c).DataBodyRange)
            sample_before = call_with_retry(lambda b=body: b.Cells(1, 1).Formula)
            call_with_retry(setattr, body, "Formula", formula)
            sample_after = call_with_retry(lambda b=body: b.Cells(1, 1).Formula)
            result["changes"].append({"sheet": "ROE_wk", "column": col_name, "sample_before": str(sample_before), "sample_after": str(sample_after)})

        # Force pivot refresh after formula-source change.
        pivot_errors = []
        for ws in wb.Worksheets:
            try:
                pts = ws.PivotTables()
                for i in range(1, pts.Count + 1):
                    pt = pts.Item(i)
                    try:
                        call_with_retry(pt.RefreshTable)
                    except Exception as exc:  # noqa: BLE001
                        pivot_errors.append({"sheet": ws.Name, "pivot": pt.Name, "error": str(exc)})
            except Exception:
                pass
        result["pivot_refresh_errors"] = pivot_errors

        call_with_retry(setattr, excel, "Calculation", -4105)  # automatic
        call_with_retry(excel.CalculateFullRebuild)
        call_with_retry(wb.Save)
        call_with_retry(wb.Close, SaveChanges=False)
        wb = None
        REPORT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"Updated source formulas and refreshed pivots: {WORKBOOK}")
        print(f"Backup: {BACKUP}")
        print(f"Report: {REPORT}")
        print(f"Pivot refresh errors: {len(pivot_errors)}")
        return 0
    except Exception as exc:  # noqa: BLE001
        print(f"Failed: {exc}")
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        return 1
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    raise SystemExit(main())
