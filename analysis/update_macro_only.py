from __future__ import annotations
import json, shutil, time
from datetime import datetime
from pathlib import Path
import pythoncom, pywintypes, win32com.client

ROOT=Path(r'C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX')
WORKBOOK=Path(r'C:\Users\VNO024\Downloads\Base_DSU2026 - TbM - WK19.xlsm')
BACKUP_DIR=ROOT/'backups'
ANALYSIS_DIR=ROOT/'analysis'
REPORT=ANALYSIS_DIR/'Base_DSU2026_WK19_update_macro_only_report.json'
PROGRESS=ANALYSIS_DIR/'Base_DSU2026_WK19_update_macro_only_progress.log'
MODULE_NAME='modCodexAtualizacaoDashboards'
MACRO_NAME='AtualizarDashboardsSemPreenchimento'

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

    ' Atualiza o cache completo primeiro: isso nao depende do filtro visivel da pivot.
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

    ' Atualiza visualizacoes. Limpa filtros temporarios, mas preserva filtros oficiais como OTO/Atrasado/Volume.
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
    print('[macro-only]', msg, flush=True)
    with PROGRESS.open('a',encoding='utf-8') as f: f.write(datetime.now().isoformat(timespec='seconds')+' '+msg+'\n')

def retry(fn,*args,retries=20,delay=.35,**kwargs):
    last=None
    for _ in range(retries):
        try: return fn(*args,**kwargs)
        except pywintypes.com_error as e:
            last=e; time.sleep(delay); pythoncom.PumpWaitingMessages()
    raise last

def main():
    if PROGRESS.exists(): PROGRESS.unlink()
    BACKUP_DIR.mkdir(exist_ok=True); ANALYSIS_DIR.mkdir(exist_ok=True)
    backup=BACKUP_DIR/f"Base_DSU2026 - TbM - WK19_before_macro_only_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsm"
    shutil.copy2(WORKBOOK, backup)
    report={'workbook':str(WORKBOOK),'backup':str(backup),'started_at':datetime.now().isoformat(timespec='seconds')}
    pythoncom.CoInitialize()
    excel=win32com.client.DispatchEx('Excel.Application')
    excel.Visible=False; excel.DisplayAlerts=False; excel.EnableEvents=False; excel.AskToUpdateLinks=False
    try:
        try: excel.AutomationSecurity=3
        except Exception: pass
        try: excel.CalculateBeforeSave=False
        except Exception: pass
        log('opening')
        wb=retry(excel.Workbooks.Open, str(WORKBOOK), UpdateLinks=False, ReadOnly=False, IgnoreReadOnlyRecommended=True, AddToMru=False, Notify=False)
        try:
            log('opened')
            report['opened']=True
            vb=wb.VBProject
            comp=None
            for i in range(1,int(vb.VBComponents.Count)+1):
                c=vb.VBComponents.Item(i)
                if c.Name=='modCodexAtualizacaoDashboards': comp=c; break
            if comp is None:
                comp=vb.VBComponents.Add(1); comp.Name='modCodexAtualizacaoDashboards'; report['module_created']=True
            else:
                report['module_created']=False
            cm=comp.CodeModule
            if cm.CountOfLines: cm.DeleteLines(1, cm.CountOfLines)
            cm.AddFromString(VBA_CODE)
            report['module_lines']=int(cm.CountOfLines)
            log('saving')
            retry(wb.Save)
            log('saved')
            report['saved']=True
            # validate module text
            report['macro_found_after']='AtualizarDashboardsSemPreenchimento' in cm.Lines(1, cm.CountOfLines)
            names=[str(wb.Worksheets.Item(i).Name) for i in range(1,int(wb.Worksheets.Count)+1)]
            report['region_errors_exists']= 'Region errors' in names
        finally:
            log('closing')
            wb.Close(SaveChanges=False)
    finally:
        try: excel.Quit()
        except Exception: pass
        pythoncom.CoUninitialize()
    report['finished_at']=datetime.now().isoformat(timespec='seconds')
    report['workbook_size']=WORKBOOK.stat().st_size
    report['workbook_mtime']=datetime.fromtimestamp(WORKBOOK.stat().st_mtime).isoformat(timespec='seconds')
    REPORT.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(report,ensure_ascii=False,indent=2))
    print('REPORT='+str(REPORT))
if __name__=='__main__': main()
