from __future__ import annotations

import json, shutil, time, re
from datetime import datetime
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

import pythoncom, pywintypes, win32com.client

ROOT=Path(r'C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX')
WORKBOOK=Path(r'C:\Users\VNO024\Downloads\Base_DSU2026 - TbM - WK19.xlsm')
BACKUP_DIR=ROOT/'backups'
ANALYSIS_DIR=ROOT/'analysis'
REPORT=ANALYSIS_DIR/'Base_DSU2026_WK19_region_errors_fast_apply_report.json'
PROGRESS=ANALYSIS_DIR/'Base_DSU2026_WK19_region_errors_fast_apply_progress.log'
OLD_SHEET='Sem_Preenchimento_Regiao'
NEW_SHEET='Region errors'
MODULE_NAME='modCodexAtualizacaoDashboards'
MACRO_NAME='AtualizarDashboardsSemPreenchimento'
BUTTON_NAME='btnAtualizarDashboardsSemPreenchimento'
BUTTON_TEXT='Atualizar Dashboards'
BUTTON_ANCHOR='K2'
REGIONS=['North','Northeast','Southeast','South','Sem Porto/Region']
DAYS=['Sun','Mon','Tue','Wed','Thu','Fri','Sat']
TARGET_SHEETS=['Volume_DS','Volume_MAO','Volume_Graph','Week_Overview','Top_Offenders_Customers','Top_Offenders_Vendors']
XL_CALC_MANUAL=-4135
XL_CALC_AUTOMATIC=-4105

VBA_CODE = r'''
Option Explicit

Public Sub AtualizarDashboardsSemPreenchimento(Optional ByVal showMessage As Boolean = True)
    Dim targetSheets As Variant
    Dim sheetName As Variant
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim co As ChartObject
    Dim refreshedCaches As Object
    Dim cacheKey As String
    Dim activeWeek As Variant
    Dim updatedPivots As Long
    Dim updatedCharts As Long
    Dim updatedSheets As Long
    Dim updatedCaches As Long
    Dim resetFilters As Long
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldStatusBar As Variant
    Dim oldCalculation As XlCalculation
    Dim errors As String

    On Error GoTo FalhaGeral

    targetSheets = Array("Volume_DS", "Volume_MAO", "Volume_Graph", "Week_Overview", "Top_Offenders_Customers", "Top_Offenders_Vendors")
    Set refreshedCaches = CreateObject("Scripting.Dictionary")
    activeWeek = ThisWorkbook.Worksheets("Week_Overview").Range("AG1").Value

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldStatusBar = Application.StatusBar
    oldCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Atualizando dashboards DSU..."

    ' 1) Atualiza os caches primeiro. Isso atualiza a base inteira da pivot, nao apenas o filtro visivel.
    For Each sheetName In targetSheets
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
        For Each pt In ws.PivotTables
            On Error Resume Next
            cacheKey = CStr(pt.CacheIndex)
            If Not refreshedCaches.Exists(cacheKey) Then
                pt.PivotCache.MissingItemsLimit = xlMissingItemsNone
                pt.PivotCache.Refresh
                If Err.Number <> 0 Then
                    errors = errors & vbCrLf & CStr(sheetName) & ": erro no cache da pivot '" & pt.Name & "' - " & Err.Description
                    Err.Clear
                Else
                    refreshedCaches.Add cacheKey, True
                    updatedCaches = updatedCaches + 1
                End If
            End If
            On Error GoTo FalhaGeral
        Next pt
    Next sheetName

    ' 2) Depois atualiza cada pivot/grafico e limpa filtros temporarios de visualizacao.
    For Each sheetName In targetSheets
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
        updatedSheets = updatedSheets + 1

        For Each pt In ws.PivotTables
            On Error Resume Next
            pt.ManualUpdate = True
            pt.PreserveFormatting = True
            resetFilters = resetFilters + LimparFiltrosTemporarios(pt, activeWeek)
            pt.ManualUpdate = False
            pt.RefreshTable
            If Err.Number <> 0 Then
                errors = errors & vbCrLf & CStr(sheetName) & ": erro na pivot '" & pt.Name & "' - " & Err.Description
                Err.Clear
                pt.ManualUpdate = False
            Else
                updatedPivots = updatedPivots + 1
            End If
            On Error GoTo FalhaGeral
        Next pt

        On Error Resume Next
        ws.Calculate
        If Err.Number <> 0 Then
            errors = errors & vbCrLf & CStr(sheetName) & ": erro ao recalcular aba - " & Err.Description
            Err.Clear
        End If
        On Error GoTo FalhaGeral

        For Each co In ws.ChartObjects
            On Error Resume Next
            co.Chart.Refresh
            If Err.Number <> 0 Then
                errors = errors & vbCrLf & CStr(sheetName) & ": erro no grafico '" & co.Name & "' - " & Err.Description
                Err.Clear
            Else
                updatedCharts = updatedCharts + 1
            End If
            On Error GoTo FalhaGeral
        Next co
    Next sheetName

    On Error Resume Next
    ThisWorkbook.Worksheets("Region errors").Calculate
    On Error GoTo FalhaGeral

    ThisWorkbook.Save

Finalizar:
    Application.StatusBar = oldStatusBar
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    Application.Calculation = oldCalculation

    If showMessage Then
        If Len(errors) > 0 Then
            MsgBox "Atualizacao concluida com avisos." & vbCrLf & _
                   "Abas: " & updatedSheets & vbCrLf & _
                   "Caches: " & updatedCaches & vbCrLf & _
                   "Pivots: " & updatedPivots & vbCrLf & _
                   "Graficos: " & updatedCharts & vbCrLf & _
                   "Filtros temporarios limpos: " & resetFilters & vbCrLf & vbCrLf & _
                   errors, vbExclamation, "Atualizar dashboards"
        Else
            MsgBox "Dashboards atualizados com sucesso." & vbCrLf & _
                   "Abas: " & updatedSheets & vbCrLf & _
                   "Caches: " & updatedCaches & vbCrLf & _
                   "Pivots: " & updatedPivots & vbCrLf & _
                   "Graficos: " & updatedCharts & vbCrLf & _
                   "Filtros temporarios limpos: " & resetFilters, vbInformation, "Atualizar dashboards"
        End If
    End If
    Exit Sub

FalhaGeral:
    errors = errors & vbCrLf & "Erro geral: " & Err.Description
    Resume Finalizar
End Sub

Private Function LimparFiltrosTemporarios(ByVal pt As PivotTable, ByVal activeWeek As Variant) As Long
    Dim pf As PivotField
    Dim n As String
    Dim resets As Long

    On Error Resume Next
    For Each pf In pt.PivotFields
        n = LCase$(pf.Name)
        If EhCampoSemana(n) Then
            pf.ClearAllFilters
            If pf.Orientation = xlPageField Then
                Err.Clear
                pf.CurrentPage = CStr(activeWeek)
                If Err.Number <> 0 Then
                    Err.Clear
                    pf.CurrentPage = CLng(activeWeek)
                End If
            End If
        ElseIf pf.Orientation = xlPageField And EhFiltroTemporario(n) Then
            pf.ClearAllFilters
            resets = resets + 1
        End If
    Next pf
    LimparFiltrosTemporarios = resets
End Function

Private Function EhCampoSemana(ByVal n As String) As Boolean
    EhCampoSemana = (InStr(1, n, "weeknum", vbTextCompare) > 0) _
        Or (InStr(1, n, "week num", vbTextCompare) > 0) _
        Or (InStr(1, n, "semana", vbTextCompare) > 0)
End Function

Private Function EhFiltroTemporario(ByVal n As String) As Boolean
    ' Limpa filtros de navegacao. Preserva filtros oficiais como OTO, Atrasado, OTD e Volume.
    EhFiltroTemporario = (InStr(1, n, "region", vbTextCompare) > 0) _
        Or (InStr(1, n, "regiao", vbTextCompare) > 0) _
        Or (InStr(1, n, "porto", vbTextCompare) > 0) _
        Or (InStr(1, n, "port", vbTextCompare) > 0) _
        Or (InStr(1, n, "cliente", vbTextCompare) > 0) _
        Or (InStr(1, n, "customer", vbTextCompare) > 0) _
        Or (InStr(1, n, "provedor", vbTextCompare) > 0) _
        Or (InStr(1, n, "provider", vbTextCompare) > 0) _
        Or (InStr(1, n, "vendor", vbTextCompare) > 0)
End Function
'''.strip()

def log(msg):
    print(f'[fast-apply] {msg}', flush=True)
    with PROGRESS.open('a', encoding='utf-8') as f:
        f.write(datetime.now().isoformat(timespec='seconds')+' '+msg+'\n')

def retry(fn,*args,retries=20,delay=.35,**kwargs):
    last=None
    for _ in range(retries):
        try: return fn(*args,**kwargs)
        except pywintypes.com_error as e:
            last=e; time.sleep(delay); pythoncom.PumpWaitingMessages()
    raise last

def backup():
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    b=BACKUP_DIR/f"Base_DSU2026 - TbM - WK19_before_region_errors_fast2_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsm"
    shutil.copy2(WORKBOOK,b)
    return b

def patch_calc_xml(path):
    tmp=path.with_suffix('.calcpatch.xlsm')
    with ZipFile(path,'r') as zin, ZipFile(tmp,'w',compression=ZIP_DEFLATED,allowZip64=True) as zout:
        for item in zin.infolist():
            data=zin.read(item.filename)
            if item.filename=='xl/workbook.xml':
                text=data.decode('utf-8',errors='replace')
                def repl(m):
                    tag=m.group(0)
                    for k,v in [('calcMode','auto'),('forceFullCalc','0'),('fullCalcOnLoad','0'),('calcOnSave','1')]:
                        if re.search(rf'\b{k}="[^"]*"', tag):
                            tag=re.sub(rf'\b{k}="[^"]*"', f'{k}="{v}"', tag)
                        else:
                            tag=tag[:-2]+f' {k}="{v}"/>' if tag.endswith('/>') else tag[:-1]+f' {k}="{v}">'
                    return tag
                if re.search(r'<calcPr\b[^>]*>', text): text=re.sub(r'<calcPr\b[^>]*>', repl, text, count=1)
                data=text.encode('utf-8')
            zout.writestr(item,data)
    tmp.replace(path)

def ensure_module(wb, report):
    log('updating VBA module')
    vb=wb.VBProject
    comp=None
    for i in range(1,int(vb.VBComponents.Count)+1):
        c=vb.VBComponents.Item(i)
        if c.Name==MODULE_NAME:
            comp=c; break
    if comp is None:
        comp=vb.VBComponents.Add(1); comp.Name=MODULE_NAME; report['module_created']=True
    else:
        report['module_created']=False
    cm=comp.CodeModule
    if cm.CountOfLines: cm.DeleteLines(1, cm.CountOfLines)
    cm.AddFromString(VBA_CODE)
    report['module_lines']=int(cm.CountOfLines)

def setup_sheet(wb, report):
    log('setup sheet')
    names=[str(wb.Worksheets.Item(i).Name) for i in range(1,int(wb.Worksheets.Count)+1)]
    if NEW_SHEET in names:
        ws=wb.Worksheets(NEW_SHEET); report['sheet_renamed']=False
    else:
        ws=wb.Worksheets(OLD_SHEET); ws.Name=NEW_SHEET; report['sheet_renamed']=True
    ws.Range('A1:L22').Clear()
    ws.Range('A1').Value='Sem Preenchimento - Region errors'
    ws.Range('A2').Value='Semana ativa'; ws.Range('B2').Formula='=Week_Overview!$AG$1'
    ws.Range('A3').Value='Regra'; ws.Range('B3').Value='Volume=Ok + OTD ajustado=Sem Preenchimento'
    ws.Range('A4').Value='Observacao'; ws.Range('B4').Value='Resumo por regiao e dia da semana da semana ativa'
    ws.Range('A6').Value='Region'
    for idx,day in enumerate(DAYS,start=2): ws.Cells(6,idx).Value=day
    ws.Range('I6').Value='Total'
    start=7
    for ridx, region in enumerate(REGIONS,start=start):
        ws.Cells(ridx,1).Value=region
        for cidx, day in enumerate(DAYS,start=2):
            header_addr=ws.Cells(6,cidx).Address
            if region=='Sem Porto/Region': crit='""'
            else: crit=f'$A{ridx}'
            ws.Cells(ridx,cidx).Formula=f'=COUNTIFS(ROE_wk[Volume],"Ok",ROE_wk[OTD ajustado],"Sem Preenchimento",ROE_wk[Weeknum],$B$2,ROE_wk[Region],{crit},ROE_wk[day week],{header_addr})'
        ws.Cells(ridx,9).Formula=f'=SUM(B{ridx}:H{ridx})'
    total=start+len(REGIONS)
    ws.Cells(total,1).Value='Total Geral'
    for cidx in range(2,10):
        letter=ws.Cells(6,cidx).Address.split('$')[1]
        ws.Cells(total,cidx).Formula=f'=SUM({letter}{start}:{letter}{total-1})'
    ws.Range('A1:I1').Merge(); ws.Range('A1:I1').Font.Bold=True; ws.Range('A1:I1').Font.Size=14
    ws.Range('A6:I6').Font.Bold=True; ws.Range(f'A{total}:I{total}').Font.Bold=True
    ws.Range(f'A6:I{total}').Borders.LineStyle=1
    # fixed widths avoids slow AutoFit
    widths={1:18,2:8,3:8,4:8,5:8,6:8,7:8,8:8,9:10,11:22}
    for col,width in widths.items(): ws.Columns(col).ColumnWidth=width
    # button
    for idx in range(int(ws.Shapes.Count),0,-1):
        sh=ws.Shapes.Item(idx)
        if str(sh.Name)==BUTTON_NAME:
            sh.Delete()
    anchor=ws.Range(BUTTON_ANCHOR)
    shape=ws.Shapes.AddShape(5, anchor.Left, anchor.Top, 180, 34)
    shape.Name=BUTTON_NAME; shape.OnAction=f"'{wb.Name}'!{MACRO_NAME}"
    try:
        shape.TextFrame2.TextRange.Text=BUTTON_TEXT
        shape.TextFrame2.TextRange.Font.Size=11
        shape.TextFrame2.TextRange.Font.Bold=True
        shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB=16777215
        shape.TextFrame2.VerticalAnchor=3
        shape.TextFrame2.TextRange.ParagraphFormat.Alignment=2
        shape.Fill.ForeColor.RGB=10498160; shape.Line.ForeColor.RGB=10498160
    except Exception:
        shape.TextFrame.Characters().Text=BUTTON_TEXT
    report['summary_range']=f'A1:I{total}'
    report['button_anchor']=BUTTON_ANCHOR
    report['button_on_action']=str(shape.OnAction)

def main():
    if PROGRESS.exists(): PROGRESS.unlink()
    report={'workbook':str(WORKBOOK),'started_at':datetime.now().isoformat(timespec='seconds')}
    report['backup']=str(backup()); log('backup created')
    pythoncom.CoInitialize()
    excel=win32com.client.DispatchEx('Excel.Application')
    excel.Visible=False; excel.DisplayAlerts=False; excel.EnableEvents=False; excel.AskToUpdateLinks=False
    try:
        try: excel.AutomationSecurity=3
        except Exception: pass
        try:
            excel.Calculation = XL_CALC_MANUAL
        except Exception as e:
            report["set_calculation_manual_before_open_error"] = str(e)
        try: excel.CalculateBeforeSave=False
        except Exception: pass
        log('opening workbook')
        wb=retry(excel.Workbooks.Open, str(WORKBOOK), UpdateLinks=False, ReadOnly=False, IgnoreReadOnlyRecommended=True, AddToMru=False, Notify=False)
        try:
            log('opened')
            report['opened']=True
            try: wb.ForceFullCalculation=False; report['force_full_calculation_false']=True
            except Exception as e: report['force_full_calculation_error']=str(e)
            setup_sheet(wb, report)
            ensure_module(wb, report)
            log('saving workbook')
            retry(wb.Save)
            report['saved']=True
            log('saved')
            names=[str(wb.Worksheets.Item(i).Name) for i in range(1,int(wb.Worksheets.Count)+1)]
            report['old_sheet_exists_after']=OLD_SHEET in names
            report['new_sheet_exists_after']=NEW_SHEET in names
            ws=wb.Worksheets(NEW_SHEET)
            report['button_found_after']=any(str(ws.Shapes.Item(i).Name)==BUTTON_NAME for i in range(1,int(ws.Shapes.Count)+1))
        finally:
            log('closing workbook')
            wb.Close(SaveChanges=False)
    finally:
        try: excel.Quit()
        except Exception: pass
        pythoncom.CoUninitialize()
    log('patch calc xml')
    patch_calc_xml(WORKBOOK)
    report['calc_xml_patched']=True
    report['finished_at']=datetime.now().isoformat(timespec='seconds')
    report['workbook_size']=WORKBOOK.stat().st_size
    report['workbook_mtime']=datetime.fromtimestamp(WORKBOOK.stat().st_mtime).isoformat(timespec='seconds')
    REPORT.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(report,ensure_ascii=False,indent=2))
    print('REPORT='+str(REPORT))

if __name__=='__main__': main()

