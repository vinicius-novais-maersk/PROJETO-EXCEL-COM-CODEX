from __future__ import annotations

import json
import shutil
import time
from datetime import datetime
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import re

import pythoncom
import pywintypes
import win32com.client

ROOT = Path(r"C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX")
WORKBOOK = Path(r"C:\Users\VNO024\Downloads\Base_DSU2026 - TbM - WK19.xlsm")
BACKUP_DIR = ROOT / "backups"
ANALYSIS_DIR = ROOT / "analysis"
REPORT = ANALYSIS_DIR / "Base_DSU2026_WK19_downloads_region_errors_update_report.json"

OLD_SHEET = "Sem_Preenchimento_Regiao"
NEW_SHEET = "Region errors"
MODULE_NAME = "modCodexAtualizacaoDashboards"
MACRO_NAME = "AtualizarDashboardsSemPreenchimento"
BUTTON_NAME = "btnAtualizarDashboardsSemPreenchimento"
BUTTON_TEXT = "Atualizar Dashboards"
BUTTON_ANCHOR = "L2"
REGIONS = ["North", "Northeast", "Southeast", "South", "Sem Porto/Region"]
TARGET_SHEETS = [NEW_SHEET, "Volume_DS", "Volume_MAO", "Volume_Graph", "Week_Overview", "Top_Offenders_Customers", "Top_Offenders_Vendors"]

XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105

VBA_CODE = r'''
Option Explicit

Private Function GetActiveWeekNumber() As Variant
    On Error GoTo Fallback
    GetActiveWeekNumber = ThisWorkbook.Worksheets("Week_Overview").Range("AG1").Value
    Exit Function
Fallback:
    GetActiveWeekNumber = Empty
End Function

Private Function IsWeekField(ByVal fieldName As String) As Boolean
    Dim n As String
    n = LCase$(fieldName)
    IsWeekField = (InStr(1, n, "week", vbTextCompare) > 0 Or _
                   InStr(1, n, "semana", vbTextCompare) > 0 Or _
                   InStr(1, n, "weeknum", vbTextCompare) > 0)
End Function

Private Sub ApplyPageFilterIfExists(ByVal pf As PivotField, ByVal wantedValue As String)
    Dim pi As PivotItem
    On Error GoTo Sair
    If pf.Orientation <> xlPageField Then Exit Sub
    If Len(wantedValue) = 0 Then Exit Sub
    For Each pi In pf.PivotItems
        If CStr(pi.Name) = wantedValue Or CStr(pi.Value) = wantedValue Then
            pf.CurrentPage = pi.Name
            Exit Sub
        End If
    Next pi
Sair:
    Err.Clear
End Sub

Private Sub ApplyFirstAvailablePageFilter(ByVal pf As PivotField, ByVal options As Variant)
    Dim optionValue As Variant
    Dim pi As PivotItem
    On Error GoTo Sair
    If pf.Orientation <> xlPageField Then Exit Sub
    For Each optionValue In options
        For Each pi In pf.PivotItems
            If CStr(pi.Name) = CStr(optionValue) Or CStr(pi.Value) = CStr(optionValue) Then
                pf.CurrentPage = pi.Name
                Exit Sub
            End If
        Next pi
    Next optionValue
Sair:
    Err.Clear
End Sub

Private Sub ResetPivotFiltersSafely(ByVal pt As PivotTable, ByVal sheetName As String, ByVal activeWeek As Variant)
    Dim pf As PivotField
    Dim fieldName As String
    On Error Resume Next
    pt.ManualUpdate = True
    For Each pf In pt.PivotFields
        If pf.Orientation <> xlDataField Then pf.ClearAllFilters
    Next pf
    For Each pf In pt.PivotFields
        fieldName = pf.Name
        If IsWeekField(fieldName) Then ApplyPageFilterIfExists pf, CStr(activeWeek)
        If InStr(1, sheetName, "Top_Offenders", vbTextCompare) > 0 Then
            If InStr(1, fieldName, "OTO Out", vbTextCompare) > 0 Then ApplyPageFilterIfExists pf, "N"
            If InStr(1, fieldName, "Atrasado", vbTextCompare) > 0 Then ApplyFirstAvailablePageFilter pf, Array("1", "Atrasado", "Sim", "S")
        End If
    Next pf
    pt.ManualUpdate = False
    On Error GoTo 0
End Sub

Public Sub AtualizarDashboardsSemPreenchimento()
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
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldStatusBar As Variant
    Dim oldCalculation As XlCalculation
    Dim errors As String

    On Error GoTo FalhaGeral

    targetSheets = Array("Region errors", "Volume_DS", "Volume_MAO", "Volume_Graph", "Week_Overview", "Top_Offenders_Customers", "Top_Offenders_Vendors")
    Set refreshedCaches = CreateObject("Scripting.Dictionary")
    activeWeek = GetActiveWeekNumber()

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

    For Each sheetName In targetSheets
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
        For Each pt In ws.PivotTables
            On Error Resume Next
            cacheKey = CStr(pt.PivotCache.Index)
            If Not refreshedCaches.Exists(cacheKey) Then
                pt.PivotCache.Refresh
                refreshedCaches.Add cacheKey, True
                If Err.Number <> 0 Then
                    errors = errors & vbCrLf & CStr(sheetName) & ": erro ao atualizar cache da pivot '" & pt.Name & "' - " & Err.Description
                    Err.Clear
                Else
                    updatedCaches = updatedCaches + 1
                End If
            End If
            On Error GoTo FalhaGeral
        Next pt
    Next sheetName

    For Each sheetName In targetSheets
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
        updatedSheets = updatedSheets + 1
        For Each pt In ws.PivotTables
            On Error Resume Next
            pt.PreserveFormatting = True
            ResetPivotFiltersSafely pt, CStr(sheetName), activeWeek
            pt.RefreshTable
            If Err.Number <> 0 Then
                errors = errors & vbCrLf & CStr(sheetName) & ": erro ao atualizar pivot '" & pt.Name & "' - " & Err.Description
                Err.Clear
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
                errors = errors & vbCrLf & CStr(sheetName) & ": erro ao atualizar grafico '" & co.Name & "' - " & Err.Description
                Err.Clear
            Else
                updatedCharts = updatedCharts + 1
            End If
            On Error GoTo FalhaGeral
        Next co
    Next sheetName

    Application.Calculate
    ThisWorkbook.Save

Finalizar:
    Application.StatusBar = oldStatusBar
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    Application.Calculation = oldCalculation

    If Len(errors) > 0 Then
        MsgBox "Atualizacao concluida com avisos." & vbCrLf & _
               "Abas atualizadas: " & updatedSheets & vbCrLf & _
               "Caches atualizados: " & updatedCaches & vbCrLf & _
               "Pivots atualizadas: " & updatedPivots & vbCrLf & _
               "Graficos atualizados: " & updatedCharts & vbCrLf & vbCrLf & _
               "Avisos:" & errors, vbExclamation, "Atualizar dashboards"
    Else
        MsgBox "Dashboards atualizados com sucesso." & vbCrLf & _
               "Abas atualizadas: " & updatedSheets & vbCrLf & _
               "Caches atualizados: " & updatedCaches & vbCrLf & _
               "Pivots atualizadas: " & updatedPivots & vbCrLf & _
               "Graficos atualizados: " & updatedCharts, vbInformation, "Atualizar dashboards"
    End If
    Exit Sub

FalhaGeral:
    errors = errors & vbCrLf & "Erro geral: " & Err.Description
    Resume Finalizar
End Sub
'''.strip()


def log(msg: str) -> None:
    print(f"[region-errors-fast] {msg}", flush=True)


def call_with_retry(func, *args, retries=20, delay=0.35, **kwargs):
    last = None
    for _ in range(retries):
        try:
            return func(*args, **kwargs)
        except pywintypes.com_error as exc:
            last = exc
            time.sleep(delay)
            pythoncom.PumpWaitingMessages()
    raise last


def backup() -> Path:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    b = BACKUP_DIR / f"Base_DSU2026 - TbM - WK19_before_region_errors_fast_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsm"
    shutil.copy2(WORKBOOK, b)
    return b


def patch_calc_mode_auto(path: Path) -> None:
    tmp = path.with_suffix('.calc_tmp.xlsm')
    with ZipFile(path, 'r') as zin, ZipFile(tmp, 'w', compression=ZIP_DEFLATED, allowZip64=True) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'xl/workbook.xml':
                text = data.decode('utf-8', errors='replace')
                if re.search(r'<calcPr\b[^>]*>', text):
                    def repl(m):
                        tag = m.group(0)
                        for k, v in [('calcMode', 'auto'), ('forceFullCalc', '0'), ('fullCalcOnLoad', '0')]:
                            if re.search(rf'\b{k}="[^"]*"', tag):
                                tag = re.sub(rf'\b{k}="[^"]*"', f'{k}="{v}"', tag)
                            else:
                                tag = tag[:-2] + f' {k}="{v}"/>' if tag.endswith('/>') else tag[:-1] + f' {k}="{v}">'
                        return tag
                    text = re.sub(r'<calcPr\b[^>]*>', repl, text, count=1)
                data = text.encode('utf-8')
            zout.writestr(item, data)
    tmp.replace(path)


def ensure_module(workbook, report):
    vbproject = workbook.VBProject
    comp = None
    for idx in range(1, vbproject.VBComponents.Count + 1):
        c = vbproject.VBComponents.Item(idx)
        if c.Name == MODULE_NAME:
            comp = c
            break
    if comp is None:
        comp = vbproject.VBComponents.Add(1)
        comp.Name = MODULE_NAME
        report['module_created'] = True
    else:
        report['module_created'] = False
    cm = comp.CodeModule
    if cm.CountOfLines:
        cm.DeleteLines(1, cm.CountOfLines)
    cm.AddFromString(VBA_CODE)
    report['module_lines'] = cm.CountOfLines


def setup_sheet_and_button(workbook, report):
    names = [workbook.Worksheets.Item(i).Name for i in range(1, workbook.Worksheets.Count + 1)]
    if NEW_SHEET in names:
        ws = workbook.Worksheets(NEW_SHEET)
        report['sheet_renamed'] = False
    else:
        ws = workbook.Worksheets(OLD_SHEET)
        ws.Name = NEW_SHEET
        report['sheet_renamed'] = True

    ws.Range('C1:K20').Clear()
    ws.Range('C1').Value = 'Sem Preenchimento por dia e região'
    ws.Range('C2').Value = 'Week'
    ws.Range('D2').Formula = '=Week_Overview!$AG$1'
    ws.Range('C3').Value = 'First date'
    ws.Range('D3').Formula = '=IFERROR(MINIFS(ROE_wk[Dia],ROE_wk[Weeknum],$D$2),"")'
    headers = ['Day', 'Date', *REGIONS, 'Total']
    for i, h in enumerate(headers, start=3):
        ws.Cells(4, i).Value = h
    for i in range(7):
        row = 5 + i
        ws.Cells(row, 3).Value = f'Day {i+1}'
        ws.Cells(row, 4).Formula = f'=IF($D$3="","",$D$3+{i})'
        for col, region in zip(range(5, 10), REGIONS):
            criterion = '""' if region == 'Sem Porto/Region' else ws.Cells(4, col).Address
            ws.Cells(row, col).Formula = (
                f'=IF($D{row}="","",COUNTIFS(ROE_wk[Weeknum],$D$2,ROE_wk[Volume],"Ok",'
                f'ROE_wk[OTO Out],"N",ROE_wk[OTD ajustado],"Sem Preenchimento",ROE_wk[Dia],$D{row},ROE_wk[Region],{criterion}))'
            )
        ws.Cells(row, 10).Formula = f'=IF($D{row}="","",SUM(E{row}:I{row}))'
    ws.Range('C12').Value = 'Total'
    for col in range(5, 11):
        letter = ws.Cells(4, col).Address.split('$')[1]
        ws.Cells(12, col).Formula = f'=SUM({letter}5:{letter}11)'
    ws.Range('C1:J1').Merge()
    ws.Range('C1:J1').Font.Bold = True
    ws.Range('C4:J4').Font.Bold = True
    ws.Range('C12:J12').Font.Bold = True
    ws.Range('D5:D11').NumberFormat = 'ddd dd/mm'
    ws.Range('C4:J12').Borders.LineStyle = 1
    # Fixed widths instead of AutoFit to avoid slow workbook-wide recalculation.
    widths = {3: 10, 4: 12, 5: 12, 6: 14, 7: 14, 8: 12, 9: 18, 10: 10, 12: 20}
    for col, width in widths.items():
        ws.Columns(col).ColumnWidth = width

    for idx in range(ws.Shapes.Count, 0, -1):
        sh = ws.Shapes.Item(idx)
        if sh.Name == BUTTON_NAME:
            sh.Delete()
    anchor = ws.Range(BUTTON_ANCHOR)
    shape = ws.Shapes.AddShape(5, anchor.Left, anchor.Top, 170, 34)
    shape.Name = BUTTON_NAME
    shape.OnAction = f"'{workbook.Name}'!{MACRO_NAME}"
    try:
        shape.TextFrame2.TextRange.Text = BUTTON_TEXT
        shape.TextFrame2.TextRange.Font.Size = 11
        shape.TextFrame2.TextRange.Font.Bold = True
        shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215
        shape.TextFrame2.VerticalAnchor = 3
        shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
        shape.Fill.ForeColor.RGB = 10498160
        shape.Line.ForeColor.RGB = 10498160
    except Exception:
        shape.TextFrame.Characters().Text = BUTTON_TEXT
    report['button_anchor'] = BUTTON_ANCHOR
    report['button_on_action'] = shape.OnAction
    report['panel_range'] = 'C1:J12'


def main():
    if not WORKBOOK.exists():
        raise FileNotFoundError(WORKBOOK)
    ANALYSIS_DIR.mkdir(parents=True, exist_ok=True)
    report = {'workbook': str(WORKBOOK), 'started_at': datetime.now().isoformat(timespec='seconds'), 'target_sheets': TARGET_SHEETS}
    report['backup'] = str(backup())
    log(f"Backup criado: {report['backup']}")

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    excel.AskToUpdateLinks = False
    excel.CalculateBeforeSave = False
    workbook = None
    try:
        excel.Calculation = XL_CALC_MANUAL
        workbook = call_with_retry(excel.Workbooks.Open, str(WORKBOOK), UpdateLinks=False, ReadOnly=False, IgnoreReadOnlyRecommended=True, AddToMru=False, Notify=False)
        report['opened'] = True
        try:
            workbook.ForceFullCalculation = False
            report['force_full_calculation_set_false'] = True
        except Exception as exc:
            report['force_full_calculation_error'] = str(exc)
        setup_sheet_and_button(workbook, report)
        ensure_module(workbook, report)
        call_with_retry(workbook.Save)
        report['saved'] = True
        # Validate structure without heavy recalculation.
        names = [workbook.Worksheets.Item(i).Name for i in range(1, workbook.Worksheets.Count + 1)]
        report['old_sheet_exists_after'] = OLD_SHEET in names
        report['new_sheet_exists_after'] = NEW_SHEET in names
        report['missing_target_sheets_after'] = [s for s in TARGET_SHEETS if s not in names]
        ws = workbook.Worksheets(NEW_SHEET)
        report['button_found_after'] = any(ws.Shapes.Item(i).Name == BUTTON_NAME for i in range(1, ws.Shapes.Count + 1))
    finally:
        if workbook is not None:
            try: workbook.Close(SaveChanges=False)
            except Exception: pass
        try: excel.Quit()
        except Exception: pass
        pythoncom.CoUninitialize()

    patch_calc_mode_auto(WORKBOOK)
    report['calc_mode_patched_auto'] = True
    report['finished_at'] = datetime.now().isoformat(timespec='seconds')
    report['workbook_size'] = WORKBOOK.stat().st_size
    report['workbook_mtime'] = datetime.fromtimestamp(WORKBOOK.stat().st_mtime).isoformat(timespec='seconds')
    REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(report, ensure_ascii=False, indent=2))
    print(f'REPORT={REPORT}')

if __name__ == '__main__':
    main()
