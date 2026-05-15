from __future__ import annotations
from pathlib import Path
import json, sys
import pythoncom
import win32com.client
try:
    from scripts.update_special_clients import call_with_retry
except ModuleNotFoundError:
    sys.path.insert(0, str(Path.cwd()))
    from scripts.update_special_clients import call_with_retry

WB_PATH = Path(r"C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX\analysis\Base_DSU2026_WK19_rules_hypercare_fix_20260508_173020.xlsm")
REPORT = Path(r"C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX\analysis\Base_DSU2026_WK19_aux_error_cleanup_report.json")
XL_UP = -4162


def last_row(ws, col=1):
    return int(call_with_retry(lambda: ws.Cells(ws.Rows.Count, col).End(XL_UP).Row))


def set_range_formula(ws, addr, formula, result):
    call_with_retry(setattr, ws.Range(addr), "Formula", formula)
    result["changes"].append({"sheet": ws.Name, "range": addr, "formula": formula})


def set_range_formula2(ws, addr, formula, result):
    try:
        call_with_retry(setattr, ws.Range(addr), "Formula2", formula)
    except Exception:
        call_with_retry(setattr, ws.Range(addr), "Formula", formula)
    result["changes"].append({"sheet": ws.Name, "range": addr, "formula": formula})


def main() -> int:
    result = {"workbook": str(WB_PATH), "changes": [], "pivot_refresh": [], "errors": []}
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    wb = None
    try:
        wb = call_with_retry(excel.Workbooks.Open, str(WB_PATH), UpdateLinks=False, ReadOnly=False)
        call_with_retry(setattr, excel, "Calculation", -4135)

        ws = wb.Worksheets("ROE_wk")
        lr = last_row(ws, 1)
        set_range_formula(ws, f"BH2:BH{lr}", '=IFERROR(BG2*24*60,"")', result)

        ws = wb.Worksheets("ROE_wk_cancel")
        lr = last_row(ws, 1)
        set_range_formula(ws, f"BF2:BF{lr}", '=IFERROR(BE2*24*60,"")', result)

        ws = wb.Worksheets("ROE_wk_monthly")
        lr = last_row(ws, 1)
        set_range_formula(ws, f"BF2:BF{lr}", '=IFERROR(BE2*24*60,"")', result)
        set_range_formula2(ws, f"BH2:BH{lr}", '=IF(XLOOKUP(A2,SIL_wk_month!B:B,SIL_wk_month!M:M,"")="","N","S")', result)

        ws = wb.Worksheets("Reagendas")
        lr = last_row(ws, 1)
        set_range_formula2(ws, f"P2:P{lr}", '=IFERROR(TEXTJOIN(CHAR(10), TRUE, FILTER(TEXTSPLIT(G2,CHAR(10)), TEXTSPLIT(G2,CHAR(10))<>"")),"")', result)
        set_range_formula(ws, f"Q2:Q{lr}", '=IF(P2="","",LEN(P2)-LEN(SUBSTITUTE(P2,CHAR(10),""))+1)', result)

        ws = wb.Worksheets("Query1")
        try:
            call_with_retry(setattr, ws.Range("BF2"), "Value", "")
            result["changes"].append({"sheet": "Query1", "range": "BF2", "value": ""})
        except Exception as exc:
            result["errors"].append(f"Query1 BF2 cleanup failed: {exc}")

        call_with_retry(setattr, excel, "Calculation", -4105)
        call_with_retry(excel.CalculateFullRebuild)
        try:
            call_with_retry(excel.CalculateUntilAsyncQueriesDone)
        except Exception:
            pass

        # Refresh pivot tables after source calculation. Keep failures as evidence, do not abort.
        for ws in wb.Worksheets:
            try:
                pts = ws.PivotTables()
                for i in range(1, pts.Count + 1):
                    pt = pts.Item(i)
                    try:
                        ok = call_with_retry(pt.RefreshTable)
                        result["pivot_refresh"].append({"sheet": ws.Name, "pivot": pt.Name, "ok": bool(ok)})
                    except Exception as exc:
                        result["pivot_refresh"].append({"sheet": ws.Name, "pivot": pt.Name, "ok": False, "error": str(exc)})
            except Exception:
                continue

        call_with_retry(excel.CalculateFullRebuild)
        call_with_retry(wb.Save)
        call_with_retry(wb.Close, SaveChanges=False)
        wb = None
        REPORT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps({"saved": str(WB_PATH), "changes": len(result["changes"]), "pivot_refresh": len(result["pivot_refresh"]), "errors": result["errors"], "report": str(REPORT)}, ensure_ascii=False, indent=2))
        return 0
    except Exception as exc:
        result["errors"].append(str(exc))
        REPORT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps(result, ensure_ascii=False, indent=2))
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
