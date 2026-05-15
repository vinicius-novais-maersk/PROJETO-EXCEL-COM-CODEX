from __future__ import annotations

import json
import shutil
import time
from datetime import datetime
from pathlib import Path

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
BUTTON_ANCHOR = "K2"
TARGET_SHEETS = [
    "Volume_DS",
    "Volume_MAO",
    "Volume_Graph",
    "Week_Overview",
    "Top_Offenders_Customers",
    "Top_Offenders_Vendors",
]
DAY_HEADERS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
REGIONS = ["North", "Northeast", "Southeast", "South", "Sem Porto/Region"]

VBA_CODE = r'''
Option Explicit

Public Sub AtualizarDashboardsSemPreenchimento(Optional ByVal showMessage As Boolean = True)
    Dim targetSheets As Variant
    Dim sheetName As Variant
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim co As ChartObject
    Dim cacheKey As String
    Dim refreshedCaches As Object
    Dim activeWeek As Variant
    Dim updatedPivots As Long
    Dim updatedCharts As Long
    Dim updatedSheets As Long
    Dim refreshedCacheCount As Long
    Dim resetFilterCount As Long
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldStatusBar As Variant
    Dim oldCalculation As XlCalculation
    Dim errors As String

    On Error GoTo FalhaGeral

    targetSheets = Array( _
        "Volume_DS", _
        "Volume_MAO", _
        "Volume_Graph", _
        "Week_Overview", _
        "Top_Offenders_Customers", _
        "Top_Offenders_Vendors" _
    )

    activeWeek = ThisWorkbook.Worksheets("Week_Overview").Range("AG1").Value
    Set refreshedCaches = CreateObject("Scripting.Dictionary")

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
        updatedSheets = updatedSheets + 1

        For Each pt In ws.PivotTables
            cacheKey = CStr(pt.CacheIndex)

            If Not refreshedCaches.Exists(cacheKey) Then
                On Error Resume Next
                pt.PivotCache.MissingItemsLimit = xlMissingItemsNone
                pt.PivotCache.Refresh
                If Err.Number <> 0 Then
                    errors = errors & vbCrLf & CStr(sheetName) & ": erro ao atualizar cache da pivot '" & pt.Name & "' - " & Err.Description
                    Err.Clear
                Else
                    refreshedCaches.Add cacheKey, True
                    refreshedCacheCount = refreshedCacheCount + 1
                End If
                On Error GoTo FalhaGeral
            End If

            On Error Resume Next
            pt.ManualUpdate = True
            pt.PreserveFormatting = True
            resetFilterCount = resetFilterCount + AplicarFiltrosOficiaisPivot(pt, activeWeek)
            pt.ManualUpdate = False
            pt.RefreshTable
            If Err.Number <> 0 Then
                errors = errors & vbCrLf & CStr(sheetName) & ": erro ao atualizar pivot '" & pt.Name & "' - " & Err.Description
                Err.Clear
                pt.ManualUpdate = False
            Else
                updatedPivots = updatedPivots + 1
            End If
            On Error GoTo FalhaGeral
        Next pt

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

        On Error Resume Next
        ws.Calculate
        If Err.Number <> 0 Then
            errors = errors & vbCrLf & CStr(sheetName) & ": erro ao recalcular aba - " & Err.Description
            Err.Clear
        End If
        On Error GoTo FalhaGeral
    Next sheetName

    On Error Resume Next
    ThisWorkbook.Worksheets("Region errors").Calculate
    If Err.Number <> 0 Then
        errors = errors & vbCrLf & "Region errors: erro ao recalcular aba - " & Err.Description
        Err.Clear
    End If
    On Error GoTo FalhaGeral

    Application.Calculate
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
                   "Abas atualizadas: " & updatedSheets & vbCrLf & _
                   "Caches de pivot atualizados: " & refreshedCacheCount & vbCrLf & _
                   "Pivots atualizadas: " & updatedPivots & vbCrLf & _
                   "Graficos atualizados: " & updatedCharts & vbCrLf & _
                   "Filtros de usuario limpos: " & resetFilterCount & vbCrLf & vbCrLf & _
                   "Avisos:" & errors, vbExclamation, "Atualizar dashboards"
        Else
            MsgBox "Dashboards atualizados com sucesso." & vbCrLf & _
                   "Abas atualizadas: " & updatedSheets & vbCrLf & _
                   "Caches de pivot atualizados: " & refreshedCacheCount & vbCrLf & _
                   "Pivots atualizadas: " & updatedPivots & vbCrLf & _
                   "Graficos atualizados: " & updatedCharts & vbCrLf & _
                   "Filtros de usuario limpos: " & resetFilterCount, vbInformation, "Atualizar dashboards"
        End If
    End If
    Exit Sub

FalhaGeral:
    errors = errors & vbCrLf & "Erro geral: " & Err.Description
    Resume Finalizar
End Sub

Private Function AplicarFiltrosOficiaisPivot(ByVal pt As PivotTable, ByVal activeWeek As Variant) As Long
    Dim pf As PivotField
    Dim fieldName As String
    Dim resetCount As Long

    On Error Resume Next

    For Each pf In pt.PivotFields
        fieldName = LCase$(pf.Name)

        If EhCampoSemana(fieldName) Then
            pf.ClearAllFilters
            If pf.Orientation = xlPageField Then
                Err.Clear
                pf.CurrentPage = CStr(activeWeek)
                If Err.Number <> 0 Then
                    Err.Clear
                    pf.CurrentPage = CLng(activeWeek)
                End If
            End If
        ElseIf pf.Orientation = xlPageField And EhFiltroUsuario(fieldName) Then
            pf.ClearAllFilters
            resetCount = resetCount + 1
        End If
    Next pf

    AplicarFiltrosOficiaisPivot = resetCount
End Function

Private Function EhCampoSemana(ByVal fieldName As String) As Boolean
    EhCampoSemana = (InStr(1, fieldName, "weeknum", vbTextCompare) > 0) _
        Or (InStr(1, fieldName, "week num", vbTextCompare) > 0) _
        Or (InStr(1, fieldName, "semana", vbTextCompare) > 0)
End Function

Private Function EhFiltroUsuario(ByVal fieldName As String) As Boolean
    ' Campos que normalmente sao filtros temporarios de visualizacao.
    ' Os filtros oficiais de negocio, como OTO, Atrasado, OTD e Volume, sao preservados.
    EhFiltroUsuario = (InStr(1, fieldName, "region", vbTextCompare) > 0) _
        Or (InStr(1, fieldName, "regiao", vbTextCompare) > 0) _
        Or (InStr(1, fieldName, "porto", vbTextCompare) > 0) _
        Or (InStr(1, fieldName, "port", vbTextCompare) > 0) _
        Or (InStr(1, fieldName, "cliente", vbTextCompare) > 0) _
        Or (InStr(1, fieldName, "customer", vbTextCompare) > 0) _
        Or (InStr(1, fieldName, "provedor", vbTextCompare) > 0) _
        Or (InStr(1, fieldName, "provider", vbTextCompare) > 0) _
        Or (InStr(1, fieldName, "vendor", vbTextCompare) > 0)
End Function
'''.strip()


def log(message: str) -> None:
    print(f"[region-errors] {message}", flush=True)


def call_with_retry(func, *args, retries: int = 30, delay: float = 0.35, **kwargs):
    last_exc = None
    for _ in range(retries):
        try:
            return func(*args, **kwargs)
        except pywintypes.com_error as exc:
            last_exc = exc
            time.sleep(delay)
            pythoncom.PumpWaitingMessages()
    raise last_exc


def backup_workbook() -> Path:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = BACKUP_DIR / f"Base_DSU2026 - TbM - WK19_before_region_errors_{timestamp}.xlsm"
    shutil.copy2(WORKBOOK, backup)
    return backup


def sheet_names(workbook) -> list[str]:
    return [str(workbook.Worksheets.Item(i).Name) for i in range(1, int(workbook.Worksheets.Count) + 1)]


def ensure_region_errors_sheet(workbook, report: dict):
    names = sheet_names(workbook)
    if NEW_SHEET in names:
        ws = workbook.Worksheets(NEW_SHEET)
        report["sheet_renamed"] = False
    elif OLD_SHEET in names:
        ws = workbook.Worksheets(OLD_SHEET)
        ws.Name = NEW_SHEET
        report["sheet_renamed"] = True
    else:
        ws = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
        ws.Name = NEW_SHEET
        report["sheet_created"] = True
    return ws


def write_region_day_summary(ws, report: dict) -> None:
    # Clear only the monitoring area. Keep the rest of the workbook untouched.
    ws.Range("A1:L25").Clear()

    ws.Range("A1").Value = "Sem Preenchimento - Region errors"
    ws.Range("A2").Value = "Semana ativa"
    ws.Range("B2").Formula = "=Week_Overview!$AG$1"
    ws.Range("A3").Value = "Regra"
    ws.Range("B3").Value = "Volume=Ok + OTD ajustado=Sem Preenchimento; fora do KPI OTO"
    ws.Range("A4").Value = "Atualizar"
    ws.Range("B4").Value = "Use o botão à direita para atualizar dashboards, pivots e gráficos"

    ws.Range("A6").Value = "Region"
    for idx, day in enumerate(DAY_HEADERS, start=2):
        ws.Cells(6, idx).Value = day
    ws.Range("I6").Value = "Total"

    start_row = 7
    for ridx, region in enumerate(REGIONS, start=start_row):
        ws.Cells(ridx, 1).Value = region
        for cidx, day in enumerate(DAY_HEADERS, start=2):
            if region == "Sem Porto/Region":
                formula = f'=COUNTIFS(ROE_wk[Volume],"Ok",ROE_wk[OTD ajustado],"Sem Preenchimento",ROE_wk[Weeknum],$B$2,ROE_wk[Region],"",ROE_wk[day week],{ws.Cells(6, cidx).Address})'
            else:
                formula = f'=COUNTIFS(ROE_wk[Volume],"Ok",ROE_wk[OTD ajustado],"Sem Preenchimento",ROE_wk[Weeknum],$B$2,ROE_wk[Region],$A{ridx},ROE_wk[day week],{ws.Cells(6, cidx).Address})'
            ws.Cells(ridx, cidx).Formula = formula
        ws.Cells(ridx, 9).Formula = f"=SUM(B{ridx}:H{ridx})"

    total_row = start_row + len(REGIONS)
    ws.Cells(total_row, 1).Value = "Total Geral"
    for cidx in range(2, 10):
        col_letter = ws.Cells(6, cidx).Address.split("$")[1]
        ws.Cells(total_row, cidx).Formula = f"=SUM({col_letter}{start_row}:{col_letter}{total_row-1})"

    # Basic formatting.
    title = ws.Range("A1:I1")
    title.Merge()
    title.Font.Bold = True
    title.Font.Size = 14
    title.Interior.Color = 15123099

    header = ws.Range("A6:I6")
    header.Font.Bold = True
    header.Interior.Color = 15123099
    header.Borders.LineStyle = 1

    body = ws.Range(f"A7:I{total_row}")
    body.Borders.LineStyle = 1
    ws.Range(f"A{total_row}:I{total_row}").Font.Bold = True
    ws.Range(f"A{total_row}:I{total_row}").Interior.Color = 13434879

    ws.Columns("A:I").AutoFit()
    ws.Columns("B:I").HorizontalAlignment = -4108  # xlCenter
    ws.Range("B2").HorizontalAlignment = -4108

    report["summary_range"] = f"A1:I{total_row}"
    report["summary_regions"] = REGIONS
    report["summary_days"] = DAY_HEADERS


def ensure_vba_module(workbook, report: dict) -> None:
    try:
        vbproject = workbook.VBProject
    except Exception as exc:
        raise RuntimeError(
            "Nao consegui acessar o VBA Project. Ative no Excel: File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model."
        ) from exc

    component = None
    for idx in range(1, int(vbproject.VBComponents.Count) + 1):
        candidate = vbproject.VBComponents.Item(idx)
        if candidate.Name == MODULE_NAME:
            component = candidate
            break

    if component is None:
        component = vbproject.VBComponents.Add(1)
        component.Name = MODULE_NAME
        report["module_created"] = True
    else:
        report["module_created"] = False

    code_module = component.CodeModule
    existing_lines = int(code_module.CountOfLines)
    if existing_lines:
        code_module.DeleteLines(1, existing_lines)
    code_module.AddFromString(VBA_CODE)
    report["module_name"] = MODULE_NAME
    report["macro_name"] = MACRO_NAME
    report["module_lines"] = int(code_module.CountOfLines)


def remove_existing_buttons(ws) -> int:
    removed = 0
    for idx in range(int(ws.Shapes.Count), 0, -1):
        shape = ws.Shapes.Item(idx)
        if str(shape.Name) == BUTTON_NAME or str(getattr(shape, "OnAction", "")).endswith(MACRO_NAME):
            shape.Delete()
            removed += 1
    return removed


def add_button(workbook, ws, report: dict) -> None:
    report["existing_buttons_removed"] = remove_existing_buttons(ws)
    anchor = ws.Range(BUTTON_ANCHOR)
    shape = ws.Shapes.AddShape(5, anchor.Left, anchor.Top, 180, 36)
    shape.Name = BUTTON_NAME
    shape.OnAction = f"'{workbook.Name}'!{MACRO_NAME}"
    try:
        shape.TextFrame2.TextRange.Text = BUTTON_TEXT
        shape.TextFrame2.TextRange.Font.Size = 11
        shape.TextFrame2.TextRange.Font.Bold = True
        shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215
        shape.TextFrame2.VerticalAnchor = 3
        shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    except Exception:
        shape.TextFrame.Characters().Text = BUTTON_TEXT
    try:
        shape.Fill.ForeColor.RGB = 10498160
        shape.Line.ForeColor.RGB = 10498160
    except Exception:
        pass
    report["button_name"] = BUTTON_NAME
    report["button_text"] = BUTTON_TEXT
    report["button_anchor"] = BUTTON_ANCHOR
    report["button_on_action"] = shape.OnAction


def read_summary_values(ws) -> list[list[object]]:
    values = []
    for r in range(6, 13):
        row = []
        for c in range(1, 10):
            row.append(ws.Cells(r, c).Value)
        values.append(row)
    return values


def validate(workbook, ws, report: dict) -> None:
    names = sheet_names(workbook)
    report["old_sheet_exists_after"] = OLD_SHEET in names
    report["new_sheet_exists_after"] = NEW_SHEET in names
    report["missing_target_sheets_after"] = [s for s in TARGET_SHEETS if s not in names]

    found_button = False
    button_action = None
    for idx in range(1, int(ws.Shapes.Count) + 1):
        shape = ws.Shapes.Item(idx)
        if str(shape.Name) == BUTTON_NAME:
            found_button = True
            button_action = str(shape.OnAction)
            break
    report["button_found_after"] = found_button
    report["button_on_action_after"] = button_action

    module_found = False
    macro_found = False
    try:
        vbproject = workbook.VBProject
        for idx in range(1, int(vbproject.VBComponents.Count) + 1):
            comp = vbproject.VBComponents.Item(idx)
            if comp.Name == MODULE_NAME:
                module_found = True
                text = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
                macro_found = f"Sub {MACRO_NAME}" in text or f"Sub {MACRO_NAME}(" in text
                break
    except Exception as exc:
        report["vba_validation_error"] = str(exc)
    report["module_found_after"] = module_found
    report["macro_found_after"] = macro_found
    report["summary_values_after_calculate"] = read_summary_values(ws)


def main() -> int:
    if not WORKBOOK.exists():
        raise FileNotFoundError(WORKBOOK)
    ANALYSIS_DIR.mkdir(parents=True, exist_ok=True)

    report = {
        "workbook": str(WORKBOOK),
        "started_at": datetime.now().isoformat(timespec="seconds"),
        "old_sheet": OLD_SHEET,
        "new_sheet": NEW_SHEET,
        "target_sheets": TARGET_SHEETS,
    }

    backup = backup_workbook()
    report["backup"] = str(backup)
    log(f"Backup criado: {backup}")

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    excel.AskToUpdateLinks = False

    workbook = None
    try:
        workbook = call_with_retry(
            excel.Workbooks.Open,
            str(WORKBOOK),
            UpdateLinks=False,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
            Notify=False,
        )
        report["opened"] = True
        report["calculation_before"] = int(excel.Calculation)
        try:
            report["force_full_calculation_before"] = bool(workbook.ForceFullCalculation)
            workbook.ForceFullCalculation = False
            report["force_full_calculation_after"] = bool(workbook.ForceFullCalculation)
        except Exception as exc:
            report["force_full_calculation_error"] = str(exc)

        names = sheet_names(workbook)
        report["missing_target_sheets_before"] = [s for s in TARGET_SHEETS if s not in names]
        if report["missing_target_sheets_before"]:
            raise RuntimeError(f"Abas ausentes: {report['missing_target_sheets_before']}")

        ws = ensure_region_errors_sheet(workbook, report)
        write_region_day_summary(ws, report)
        ensure_vba_module(workbook, report)
        add_button(workbook, ws, report)

        excel.Calculation = -4105  # xlCalculationAutomatic
        ws.Calculate()
        call_with_retry(workbook.Save)
        report["saved"] = True
        validate(workbook, ws, report)
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        try:
            excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()

    report["finished_at"] = datetime.now().isoformat(timespec="seconds")
    report["workbook_size"] = WORKBOOK.stat().st_size
    report["workbook_mtime"] = datetime.fromtimestamp(WORKBOOK.stat().st_mtime).isoformat(timespec="seconds")
    REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2, default=str))
    print(f"REPORT={REPORT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
