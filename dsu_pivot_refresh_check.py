from __future__ import annotations
import json, sys, time, traceback
from pathlib import Path
import pythoncom
import win32com.client as win32
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
WORKBOOK=Path(r'C:\Users\VNO024\OneDrive - Maersk Group\2025\Py\DSU 2025\Base_DSU2026.xlsm')
OUT=Path(r'C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX\dsu_pivot_refresh_check_result.json')

def s(x):
    try: return None if x is None else str(x)
    except Exception: return repr(x)

def main():
    res={'workbook':str(WORKBOOK),'started_at':time.strftime('%Y-%m-%d %H:%M:%S'),'pivots':[]}
    pythoncom.CoInitialize(); excel=None; wb=None
    try:
        excel=win32.DispatchEx('Excel.Application')
        excel.Visible=False; excel.DisplayAlerts=False; excel.EnableEvents=False
        try: excel.AskToUpdateLinks=False
        except Exception: pass
        try: excel.AutomationSecurity=3
        except Exception: pass
        wb=excel.Workbooks.Open(str(WORKBOOK),UpdateLinks=0,ReadOnly=True,IgnoreReadOnlyRecommended=True,AddToMru=False)
        res['opened_read_only']=bool(wb.ReadOnly)
        for i in range(1,int(wb.Worksheets.Count)+1):
            ws=wb.Worksheets(i)
            try:
                pts=ws.PivotTables()
                for j in range(1,int(pts.Count)+1):
                    pt=pts.Item(j)
                    item={'sheet':s(ws.Name),'pivot':s(pt.Name),'range':s(pt.TableRange2.Address),'ok':None,'error':None}
                    try:
                        ok=pt.RefreshTable()
                        item['ok']=s(ok)
                    except Exception as e:
                        item['error']=repr(e)
                    res['pivots'].append(item)
            except Exception as e:
                res['pivots'].append({'sheet':s(ws.Name),'collection_error':repr(e)})
        res['finished_at']=time.strftime('%Y-%m-%d %H:%M:%S')
    except Exception as e:
        res['fatal_error']=repr(e); res['traceback']=traceback.format_exc()
    finally:
        try:
            if wb is not None: wb.Close(SaveChanges=False)
        except Exception: pass
        try:
            if excel is not None: excel.Quit()
        except Exception: pass
        pythoncom.CoUninitialize()
    OUT.write_text(json.dumps(res,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
    return res
if __name__=='__main__':
    res=main()
    print('WROTE',OUT)
    errors=[p for p in res.get('pivots',[]) if p.get('error') or p.get('collection_error')]
    print(json.dumps({'fatal_error':res.get('fatal_error'),'opened_read_only':res.get('opened_read_only'),'pivot_count':len([p for p in res.get('pivots',[]) if p.get('pivot')]),'errors':errors[:20]},ensure_ascii=False,indent=2))
    if res.get('fatal_error'):
        raise SystemExit(1)
