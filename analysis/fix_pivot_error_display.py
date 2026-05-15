from __future__ import annotations
import json
from datetime import datetime
from pathlib import Path
import pythoncom
import win32com.client
try:
    from scripts.update_special_clients import call_with_retry
except ModuleNotFoundError:
    import sys
    sys.path.insert(0, str(Path.cwd()))
    from scripts.update_special_clients import call_with_retry

WORKBOOK = Path(r"C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX\analysis\Base_DSU2026_WK19_rules_hypercare_fix_20260508_173020.xlsm")
REPORT = WORKBOOK.with_name(WORKBOOK.stem + f"_pivot_error_display_fix_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")

pythoncom.CoInitialize()
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.EnableEvents = False
result={"workbook":str(WORKBOOK),"timestamp":datetime.now().isoformat(timespec='seconds'),"pivots":[]}
wb=None
try:
    wb=call_with_retry(excel.Workbooks.Open, str(WORKBOOK), UpdateLinks=False, ReadOnly=False)
    for ws_name in ["Top_Offenders_Customers", "Top_Offenders_Vendors", "Weekly_DS"]:
        try:
            ws=wb.Worksheets(ws_name)
            pts=ws.PivotTables()
            for i in range(1, pts.Count+1):
                pt=pts.Item(i)
                before_display = bool(pt.DisplayErrorString)
                before_string = str(pt.ErrorString)
                pt.DisplayErrorString = True
                pt.ErrorString = ""
                call_with_retry(pt.RefreshTable)
                result["pivots"].append({"sheet":ws_name,"pivot":pt.Name,"before_display_error":before_display,"before_error_string":before_string,"after_error_string":""})
        except Exception as exc:
            result["pivots"].append({"sheet":ws_name,"error":str(exc)})
    call_with_retry(excel.CalculateFullRebuild)
    call_with_retry(wb.Save)
    call_with_retry(wb.Close, SaveChanges=False)
    wb=None
    REPORT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f"Pivot error display fixed. Report: {REPORT}")
finally:
    if wb is not None:
        try: wb.Close(SaveChanges=False)
        except Exception: pass
    excel.Quit()
    pythoncom.CoUninitialize()
