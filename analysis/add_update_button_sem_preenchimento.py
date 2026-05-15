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
REPORT = ANALYSIS_DIR / "Base_DSU2026_WK19_downloads_update_button_report.json"

MODULE_NAME = "modCodexAtualizacaoDashboards"
MACRO_NAME = "AtualizarDashboardsSemPreenchimento"
BUTTON_NAME = "btnAtualizarDashboardsSemPreenchimento"
BUTTON_TEXT = "Atualizar Dashboards"
TARGET_SHEETS = [
    "Volume_DS",
    "Volume_MAO",
    "Volume_Graph",
    "Week_Overview",
    "Top_Offenders_Customers",
    "Top_Offenders_Vendors",
]
BUTTON_SHEET = "Sem_Preenchimento_Regiao"

VBA_CODE = r'''
Option Explicit

Public Sub AtualizarDashboardsSemPreenchimento()
    Dim targetSheets As Variant
    Dim sheetName As Variant
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim co As ChartObject
    Dim updatedPivots As Long
    Dim updatedCharts As Long
    Dim updatedSheets As Long
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

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldStatusBar = Application.StatusBar
    oldCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = "Atualizando dashboards DSU..."

    For Each sheetName In targetSheets
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
        updatedSheets = updatedSheets + 1

        On Error Resume Next
        ws.Calculate
        If Err.Number <> 0 Then
            errors = errors & vbCrLf & CStr(sheetName) & ": erro ao recalcular aba - " & Err.Description
            Err.Clear
        End If
        On Error GoTo FalhaGeral

        For Each pt In ws.PivotTables
            On Error Resume Next
            pt.PreserveFormatting = True
            pt.RefreshTable
            If Err.Number <> 0 Then
                errors = errors & vbCrLf & CStr(sheetName) & ": erro ao atualizar pivot '" & pt.Name & "' - " & Err.Description
                Err.Clear
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
    Next sheetName

    Application.CalculateFullRebuild
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
               "Pivots atualizadas: " & updatedPivots & vbCrLf & _
               "Graficos atualizados: " & updatedCharts & vbCrLf & vbCrLf & _
               "Avisos:" & errors, vbExclamation, "Atualizar dashboards"
    Else
        MsgBox "Dashboards atualizados com sucesso." & vbCrLf & _
               "Abas atualizadas: " & updatedSheets & vbCrLf & _
               "Pivots atualizadas: " & updatedPivots & vbCrLf & _
               "Graficos atualizados: " & updatedCharts, vbInformation, "Atualizar dashboards"
    End If
    Exit Sub

FalhaGeral:
    errors = errors & vbCrLf & "Erro geral: " & Err.Description
    Resume Finalizar
End Sub
'''.strip()


def log(message: str) -> None:
    print(f"[update-button] {message}", flush=True)


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
    backup = BACKUP_DIR / f"Base_DSU2026 - TbM - WK19_before_update_button_{timestamp}.xlsm"
    shutil.copy2(WORKBOOK, backup)
    return backup


def ensure_vba_module(workbook, report: dict) -> None:
    try:
        vbproject = workbook.VBProject
    except Exception as exc:
        raise RuntimeError(
            "Nao consegui acessar o VBA Project. Ative no Excel: File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model."
        ) from exc

    component = None
    for idx in range(1, vbproject.VBComponents.Count + 1):
        candidate = vbproject.VBComponents.Item(idx)
        if candidate.Name == MODULE_NAME:
            component = candidate
            break

    if component is None:
        component = vbproject.VBComponents.Add(1)  # vbext_ct_StdModule
        component.Name = MODULE_NAME
        report["module_created"] = True
    else:
        report["module_created"] = False

    code_module = component.CodeModule
    existing_lines = code_module.CountOfLines
    if existing_lines:
        code_module.DeleteLines(1, existing_lines)
    code_module.AddFromString(VBA_CODE)
    report["module_name"] = MODULE_NAME
    report["macro_name"] = MACRO_NAME
    report["module_lines"] = code_module.CountOfLines


def remove_existing_button(ws, report: dict) -> None:
    removed = 0
    for idx in range(ws.Shapes.Count, 0, -1):
        shape = ws.Shapes.Item(idx)
        if shape.Name == BUTTON_NAME:
            shape.Delete()
            removed += 1
    report["existing_buttons_removed"] = removed


def add_button(workbook, report: dict) -> None:
    ws = workbook.Worksheets(BUTTON_SHEET)
    remove_existing_button(ws, report)

    anchor = ws.Range("D2")
    left = anchor.Left
    top = anchor.Top
    width = 170
    height = 34

    # msoShapeRoundedRectangle = 5
    shape = ws.Shapes.AddShape(5, left, top, width, height)
    shape.Name = BUTTON_NAME
    shape.OnAction = f"'{workbook.Name}'!{MACRO_NAME}"

    try:
        shape.TextFrame2.TextRange.Text = BUTTON_TEXT
        shape.TextFrame2.TextRange.Font.Size = 11
        shape.TextFrame2.TextRange.Font.Bold = True
        shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215
        shape.TextFrame2.VerticalAnchor = 3  # msoAnchorMiddle
        shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 2  # msoAlignCenter
    except Exception:
        shape.TextFrame.Characters().Text = BUTTON_TEXT

    try:
        # Maersk-like blue fill, white text.
        shape.Fill.ForeColor.RGB = 10498160
        shape.Line.ForeColor.RGB = 10498160
    except Exception:
        pass

    report["button_sheet"] = BUTTON_SHEET
    report["button_name"] = BUTTON_NAME
    report["button_text"] = BUTTON_TEXT
    report["button_on_action"] = shape.OnAction
    report["button_anchor"] = "D2"


def validate(workbook, report: dict) -> None:
    sheets = [workbook.Worksheets.Item(i).Name for i in range(1, workbook.Worksheets.Count + 1)]
    missing = [s for s in [BUTTON_SHEET, *TARGET_SHEETS] if s not in sheets]
    report["missing_sheets_after"] = missing

    ws = workbook.Worksheets(BUTTON_SHEET)
    button = None
    for idx in range(1, ws.Shapes.Count + 1):
        shape = ws.Shapes.Item(idx)
        if shape.Name == BUTTON_NAME:
            button = shape
            break
    report["button_found_after"] = button is not None
    if button is not None:
        report["button_on_action_after"] = button.OnAction

    module_found = False
    macro_found = False
    try:
        vbproject = workbook.VBProject
        for idx in range(1, vbproject.VBComponents.Count + 1):
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


def main() -> int:
    if not WORKBOOK.exists():
        raise FileNotFoundError(WORKBOOK)
    ANALYSIS_DIR.mkdir(parents=True, exist_ok=True)

    report = {
        "workbook": str(WORKBOOK),
        "started_at": datetime.now().isoformat(timespec="seconds"),
        "target_sheets": TARGET_SHEETS,
        "button_sheet": BUTTON_SHEET,
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

        sheets = [workbook.Worksheets.Item(i).Name for i in range(1, workbook.Worksheets.Count + 1)]
        missing = [s for s in [BUTTON_SHEET, *TARGET_SHEETS] if s not in sheets]
        report["missing_sheets_before"] = missing
        if missing:
            raise RuntimeError(f"Abas ausentes: {missing}")

        ensure_vba_module(workbook, report)
        add_button(workbook, report)
        call_with_retry(workbook.Save)
        report["saved"] = True
        validate(workbook, report)
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
    REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    print(f"REPORT={REPORT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
