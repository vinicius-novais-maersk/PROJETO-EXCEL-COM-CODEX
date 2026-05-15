from __future__ import annotations

from pathlib import Path
import json
import sys

import pythoncom
import win32com.client

try:
    from scripts.update_special_clients import call_with_retry
except ModuleNotFoundError:
    sys.path.insert(0, str(Path.cwd()))
    from scripts.update_special_clients import call_with_retry

WORKBOOK = Path(r"C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX\analysis\Base_DSU2026_WK19_rules_hypercare_fix_20260508_173020.xlsm")
REPORT = Path(r"C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX\analysis\Base_DSU2026_WK19_top_offenders_refresh_report.json")

PROVIDER_FORMULA = '=IF(ROE_wk[[#This Row],[Tipo Serviço]]="Transporte Rodoviário",IF(COUNTIFS(ROE_wk[Provedor],ROE_wk[[#This Row],[Provedor]],ROE_wk[Tipo Serviço],ROE_wk[[#This Row],[Tipo Serviço]],ROE_wk[Porto],ROE_wk[[#This Row],[Porto]],ROE_wk[Weeknum],ROE_wk[[#This Row],[Weeknum]],ROE_wk[Volume],"Ok",ROE_wk[OTO Out],"N",ROE_wk[OTD ajustado],"<>Sem Preenchimento")=0,"",1-COUNTIFS(ROE_wk[OTD ajustado],"Atrasado",ROE_wk[Provedor],ROE_wk[[#This Row],[Provedor]],ROE_wk[Tipo Serviço],ROE_wk[[#This Row],[Tipo Serviço]],ROE_wk[Porto],ROE_wk[[#This Row],[Porto]],ROE_wk[Weeknum],ROE_wk[[#This Row],[Weeknum]],ROE_wk[Volume],"Ok",ROE_wk[OTO Out],"N",ROE_wk[Especiais],"<>Especial")/COUNTIFS(ROE_wk[Provedor],ROE_wk[[#This Row],[Provedor]],ROE_wk[Tipo Serviço],ROE_wk[[#This Row],[Tipo Serviço]],ROE_wk[Porto],ROE_wk[[#This Row],[Porto]],ROE_wk[Weeknum],ROE_wk[[#This Row],[Weeknum]],ROE_wk[Volume],"Ok",ROE_wk[OTO Out],"N",ROE_wk[OTD ajustado],"<>Sem Preenchimento"))),"")'
CLIENT_FORMULA = '=IF(ROE_wk[[#This Row],[Tipo Serviço]]="Transporte Rodoviário",IF(COUNTIFS(ROE_wk[Cliente Proposta],ROE_wk[[#This Row],[Cliente Proposta]],ROE_wk[Tipo Serviço],ROE_wk[[#This Row],[Tipo Serviço]],ROE_wk[Porto],ROE_wk[[#This Row],[Porto]],ROE_wk[Weeknum],ROE_wk[[#This Row],[Weeknum]],ROE_wk[Volume],"Ok",ROE_wk[OTO Out],"N",ROE_wk[OTD ajustado],"<>Sem Preenchimento")=0,"",1-COUNTIFS(ROE_wk[OTD ajustado],"Atrasado",ROE_wk[Cliente Proposta],ROE_wk[[#This Row],[Cliente Proposta]],ROE_wk[Tipo Serviço],ROE_wk[[#This Row],[Tipo Serviço]],ROE_wk[Porto],ROE_wk[[#This Row],[Porto]],ROE_wk[Weeknum],ROE_wk[[#This Row],[Weeknum]],ROE_wk[Volume],"Ok",ROE_wk[OTO Out],"N",ROE_wk[Especiais],"<>Especial")/COUNTIFS(ROE_wk[Cliente Proposta],ROE_wk[[#This Row],[Cliente Proposta]],ROE_wk[Tipo Serviço],ROE_wk[[#This Row],[Tipo Serviço]],ROE_wk[Porto],ROE_wk[[#This Row],[Porto]],ROE_wk[Weeknum],ROE_wk[[#This Row],[Weeknum]],ROE_wk[Volume],"Ok",ROE_wk[OTO Out],"N",ROE_wk[OTD ajustado],"<>Sem Preenchimento"))),"")'


def main() -> int:
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    result = {"workbook": str(WORKBOOK), "pivot_refresh": [], "sample_values": {}, "errors": []}
    wb = None
    try:
        wb = call_with_retry(excel.Workbooks.Open, str(WORKBOOK), UpdateLinks=False, ReadOnly=False)
        call_with_retry(setattr, excel, "Calculation", -4135)
        roe = wb.Worksheets("ROE_wk")
        table = roe.ListObjects("ROE_wk")
        call_with_retry(setattr, table.ListColumns("% OTO Provider").DataBodyRange, "Formula", PROVIDER_FORMULA)
        call_with_retry(setattr, table.ListColumns("% OTO Client").DataBodyRange, "Formula", CLIENT_FORMULA)
        call_with_retry(setattr, excel, "Calculation", -4105)
        call_with_retry(excel.CalculateFullRebuild)
        try:
            call_with_retry(wb.RefreshAll)
        except Exception as exc:
            result["errors"].append(f"RefreshAll failed: {exc}")
        try:
            # wait async query completion where available
            call_with_retry(excel.CalculateUntilAsyncQueriesDone)
        except Exception:
            pass
        for ws in wb.Worksheets:
            try:
                pts = ws.PivotTables()
                for i in range(1, pts.Count + 1):
                    pt = pts.Item(i)
                    try:
                        refreshed = call_with_retry(pt.RefreshTable)
                        result["pivot_refresh"].append({"sheet": ws.Name, "pivot": pt.Name, "ok": bool(refreshed)})
                    except Exception as exc:
                        result["pivot_refresh"].append({"sheet": ws.Name, "pivot": pt.Name, "ok": False, "error": str(exc)})
            except Exception:
                continue
        call_with_retry(excel.CalculateFullRebuild)
        # capture target values after refresh
        toc = wb.Worksheets("Top_Offenders_Customers")
        for addr in ["D12", "D30", "D50", "D127", "D167", "G27"]:
            result["sample_values"]["Top_Offenders_Customers!" + addr] = {
                "value": str(toc.Range(addr).Value),
                "text": str(toc.Range(addr).Text),
                "formula": str(toc.Range(addr).Formula),
            }
        call_with_retry(wb.Save)
        call_with_retry(wb.Close, SaveChanges=False)
        wb = None
        REPORT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps({"saved": str(WORKBOOK), "report": str(REPORT), "samples": result["sample_values"]}, ensure_ascii=False, indent=2))
        return 0
    except Exception as exc:
        result["errors"].append(str(exc))
        REPORT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"FAILED: {exc}")
        return 1
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        excel.Quit()
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    raise SystemExit(main())
