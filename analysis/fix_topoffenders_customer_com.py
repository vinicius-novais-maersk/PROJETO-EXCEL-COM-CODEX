from __future__ import annotations
import json, shutil, time
from datetime import datetime
from pathlib import Path
import pythoncom, pywintypes, win32com.client

ROOT=Path(r'C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX')
WORKBOOK=Path(r'C:\Users\VNO024\Downloads\Base_DSU2026 - TbM - WK19.xlsm')
BACKUP_DIR=ROOT/'backups'; ANALYSIS_DIR=ROOT/'analysis'
REPORT=ANALYSIS_DIR/'Base_DSU2026_WK19_topoffenders_customer_com_fix_report.json'
SHEET='Top_Offenders_Customers'

def retry(fn,*args,retries=20,delay=.35,**kwargs):
    last=None
    for _ in range(retries):
        try: return fn(*args,**kwargs)
        except pywintypes.com_error as e:
            last=e; time.sleep(delay); pythoncom.PumpWaitingMessages()
    raise last

def add_icon_set(ws, rng_addr, report):
    try:
        rng=ws.Range(rng_addr)
        try:
            rng.FormatConditions.Delete()
        except Exception:
            pass
        fc=rng.FormatConditions.AddIconSetCondition()
        # Defaults to a 3-icon set. Use numeric thresholds aligned with existing workbook logic.
        try:
            fc.IconCriteria(2).Type = 0  # xlConditionValueNumber
            fc.IconCriteria(2).Value = 0.8
            fc.IconCriteria(3).Type = 0  # xlConditionValueNumber
            fc.IconCriteria(3).Value = 0.92
        except Exception as e:
            report.setdefault('icon_threshold_warnings',[]).append({'range':rng_addr,'warning':str(e)})
        report.setdefault('icon_ranges_applied',[]).append(rng_addr)
    except Exception as e:
        report.setdefault('icon_ranges_failed',[]).append({'range':rng_addr,'error':str(e)})

def main():
    BACKUP_DIR.mkdir(exist_ok=True); ANALYSIS_DIR.mkdir(exist_ok=True)
    backup=BACKUP_DIR/f"Base_DSU2026 - TbM - WK19_before_topoffenders_customer_com_fix_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsm"
    shutil.copy2(WORKBOOK, backup)
    report={'workbook':str(WORKBOOK),'backup':str(backup),'started_at':datetime.now().isoformat(timespec='seconds')}
    pythoncom.CoInitialize()
    excel=win32com.client.DispatchEx('Excel.Application')
    excel.Visible=False; excel.DisplayAlerts=False; excel.EnableEvents=False; excel.AskToUpdateLinks=False
    try:
        try: excel.AutomationSecurity=3
        except Exception: pass
        wb=retry(excel.Workbooks.Open, str(WORKBOOK), UpdateLinks=False, ReadOnly=False, IgnoreReadOnlyRecommended=True, AddToMru=False, Notify=False)
        try:
            report['opened']=True
            ws=wb.Worksheets(SHEET)
            # 1) G510 total formula for OTO DIA.
            ws.Range('G510').Formula = '=IFERROR(-(C510/B510-1),"")'
            # 2) Hyper Care Norte mirror: blank helper rows stay blank, not zero.
            ws.Range('T7:T510').FormulaR1C1 = '=IF(R[14]C[12]="","",R[14]C[12])'
            ws.Range('U7:U510').FormulaR1C1 = '=IF(R[14]C[12]="","",R[14]C[12])'
            ws.Range('W7:W510').FormulaR1C1 = '=IF(R[14]C[11]="","",R[14]C[11])'
            ws.Range('V7:V510').FormulaR1C1 = '=IF(OR(RC[-1]="",RC[-1]=0,RC[1]=""),"",1-RC[1]/RC[-1])'
            # 3) Apply icon conditional formatting to OTO columns.
            for rng in ['D8:D510','G8:G510','P8:P167','V7:V441','AA7:AA37','AM25:AM43']:
                add_icon_set(ws, rng, report)
            # 4) Add explanatory map block in AP:AQ.
            guide=[
                ('AP1','MAPA - Top Offenders Customers'),
                ('AP2','Fonte'),('AQ2','Todas as pivots desta aba usam ROE_wk como base.'),
                ('AP3','Oficial x Diagnóstico'),('AQ3','Week_Overview é o KPI oficial agregado; Top Offenders é diagnóstico por cliente/filtro.'),
                ('AP4','Quando pode divergir'),('AQ4','Diverge quando semana, OTO Out, Atrasado, Centro de Custo, cliente, porto ou escopo ficam diferentes.'),
                ('AP5','VISÃO GERAL A:G'),('AQ5','Mostra volume, atrasos e OTO por cliente; coluna G é OTO DIA = 1 - atraso/volume.'),
                ('AP6','HYPER CARE NORTE T:W'),('AQ6','Espelho do helper AF:AH; linhas sem helper ficam em branco para não aparecer 0 falso.'),
                ('AP7','JUSTIFICATIVA Y:AA'),('AQ7','Mostra justificativas de atraso e OTO associado.'),
                ('AP8','Menu'),('AQ8','Menu!C34/E34/G34 puxam AK24; Menu!I34/K34/M34 puxam A6 via GETPIVOTDATA.'),
                ('AP9','Gráfico'),('AQ9','Chart do Menu usa AK25:AN43 desta aba.'),
                ('AP10','Regra atual'),('AQ10','Sem Preenchimento conta volume, não entra no OTO; Especial entra no KPI e não penaliza quando atrasado.'),
                ('AP11','Checklist'),('AQ11','Antes de comparar com Week_Overview, alinhar semana e filtros oficiais.'),
            ]
            for ref,val in guide:
                ws.Range(ref).Value = val
            ws.Range('AP1:AQ1').Font.Bold = True
            ws.Range('AP2:AP11').Font.Bold = True
            ws.Columns('AP').ColumnWidth = 24
            ws.Columns('AQ').ColumnWidth = 95
            # Calculate only this sheet/ranges, not the whole workbook.
            ws.Range('G510').Calculate()
            ws.Range('T7:W510').Calculate()
            report['g510_value_after']=ws.Range('G510').Value
            report['t441_w441_after']=[ws.Range(c+'441').Value for c in ['T','U','V','W']]
            report['t442_w442_after']=[ws.Range(c+'442').Value for c in ['T','U','V','W']]
            report['guide_title_after']=ws.Range('AP1').Value
            retry(wb.Save)
            report['saved']=True
        finally:
            wb.Close(SaveChanges=False)
    finally:
        try: excel.Quit()
        except Exception: pass
        pythoncom.CoUninitialize()
    report['finished_at']=datetime.now().isoformat(timespec='seconds')
    report['workbook_size']=WORKBOOK.stat().st_size
    report['workbook_mtime']=datetime.fromtimestamp(WORKBOOK.stat().st_mtime).isoformat(timespec='seconds')
    REPORT.write_text(json.dumps(report,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
    print(json.dumps(report,ensure_ascii=False,indent=2,default=str))
    print('REPORT='+str(REPORT))
if __name__=='__main__': main()
