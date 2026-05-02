from __future__ import annotations

import json
import os
import shutil
import sys
import tempfile
import time
import subprocess
from collections.abc import Iterable
from datetime import datetime
from pathlib import Path

import pythoncom
import win32com.client as win32
import win32process


XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105
XL_OPENXML_WORKBOOK_MACRO_ENABLED = 52
VBEXT_CT_STDMODULE = 1
VBEXT_CT_CLASSMODULE = 2


STD_MODULE_NAME = "BradoCleanup"
CLASS_MODULE_NAME = "CBradoCleanupAppEvents"
THISWORKBOOK_CODE = """Option Explicit

Private Sub Workbook_Open()
    EnsureBradoCleanupEvents
End Sub

Private Sub Workbook_Activate()
    EnsureBradoCleanupEvents
End Sub
"""

STD_MODULE_CODE = """Option Explicit

Public gBradoCleanupEvents As CBradoCleanupAppEvents
Public gBradoCleanupRunning As Boolean

Public Sub EnsureBradoCleanupEvents()
    If gBradoCleanupEvents Is Nothing Then
        Set gBradoCleanupEvents = New CBradoCleanupAppEvents
        Set gBradoCleanupEvents.App = Application
    End If
End Sub

Public Sub RemoveBradoRowsFromROEWk(Optional ByVal TriggeredByRefresh As Boolean = False)
    RemoveBradoRowsFromOperationalSheets TriggeredByRefresh
End Sub

Public Sub RemoveBradoRowsFromOperationalSheets(Optional ByVal TriggeredByRefresh As Boolean = False)
    If gBradoCleanupRunning Then Exit Sub

    Dim ws As Worksheet
    Dim sheetName As Variant

    On Error GoTo SafeExit
    gBradoCleanupRunning = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For Each sheetName In Array("ROE_wk", "ROE_wk_monthly")
        If WorksheetExists(CStr(sheetName)) Then
            Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
            DeleteBradoRowsFromSheet ws
        End If
    Next sheetName

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    gBradoCleanupRunning = False
End Sub

Private Sub DeleteBradoRowsFromSheet(ByVal ws As Worksheet)
    Dim colEmbarcador As Long
    Dim lastRow As Long
    Dim r As Long

    colEmbarcador = FindHeaderColumn(ws, "Embarcador")
    If colEmbarcador = 0 Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, colEmbarcador).End(xlUp).Row
    For r = lastRow To 2 Step -1
        If InStr(1, CStr(ws.Cells(r, colEmbarcador).Value), "BRADO", vbTextCompare) > 0 Then
            ws.Rows(r).Delete
        End If
    Next r
End Sub

Private Function WorksheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    WorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long
    Dim c As Long

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If Trim$(CStr(ws.Cells(1, c).Value)) = headerName Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c
End Function
"""

CLASS_MODULE_CODE = """Option Explicit

Public WithEvents App As Application
Private mCleanupPending As Boolean

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    If Not Sh Is Nothing Then
        If Sh.Name = "ROE_wk" Then mCleanupPending = True
    End If
End Sub

Private Sub App_SheetCalculate(ByVal Sh As Object)
    On Error Resume Next
    If Not Sh Is Nothing Then
        If Sh.Name = "ROE_wk" Then mCleanupPending = True
    End If
End Sub

Private Sub App_AfterCalculate()
    On Error GoTo SafeExit
    If mCleanupPending Then
        mCleanupPending = False
        RemoveBradoRowsFromROEWk True
    End If
SafeExit:
End Sub
"""


def contiguous_blocks(rows: Iterable[int]) -> list[tuple[int, int]]:
    ordered = sorted(rows)
    if not ordered:
        return []
    blocks: list[tuple[int, int]] = []
    start = ordered[0]
    end = ordered[0]
    for row in ordered[1:]:
        if row == end + 1:
            end = row
            continue
        blocks.append((start, end))
        start = end = row
    blocks.append((start, end))
    return blocks


def open_workbook(excel, path: Path):
    try:
        return excel.Workbooks.Open(str(path), ReadOnly=False, IgnoreReadOnlyRecommended=True)
    except TypeError:
        return excel.Workbooks.Open(str(path), 0, False, 5, "", "", True)


def find_header_map(ws) -> dict[str, int]:
    used = ws.UsedRange
    headers: dict[str, int] = {}
    for col in range(1, used.Columns.Count + 1):
        value = ws.Cells(1, col).Value
        if value is not None and str(value).strip():
            headers[str(value).strip()] = col
    return headers


def find_brado_rows(ws, embarcador_col: int) -> list[int]:
    used = ws.UsedRange
    rows: list[int] = []
    for row in range(2, used.Rows.Count + 1):
        value = ws.Cells(row, embarcador_col).Value
        if isinstance(value, str) and "BRADO" in value.upper():
            rows.append(row)
    return rows


def delete_brado_rows_from_sheet(ws) -> dict[str, object]:
    headers = find_header_map(ws)
    if "Embarcador" not in headers:
        return {"sheet": ws.Name, "removed_rows": 0, "row_blocks": []}

    rows_to_delete = find_brado_rows(ws, headers["Embarcador"])
    blocks = contiguous_blocks(rows_to_delete)
    for start, end in reversed(blocks):
        ws.Rows(f"{start}:{end}").Delete()
    return {"sheet": ws.Name, "removed_rows": len(rows_to_delete), "row_blocks": blocks}


def workbook_metrics(wb) -> dict[str, object]:
    week = wb.Worksheets("Week_Overview")
    ds = wb.Worksheets("Volume_DS")
    graph = wb.Worksheets("Volume_Graph")
    mao = wb.Worksheets("Volume_MAO")
    return {
        "WeekOverview_K23": week.Range("K23").Value,
        "WeekOverview_Q23": week.Range("Q23").Value,
        "VolumeDS_M9": ds.Range("M9").Value,
        "VolumeDS_M16": ds.Range("M16").Value,
        "VolumeGraph_M12": graph.Range("M12").Value,
        "VolumeGraph_M17": graph.Range("M17").Value,
        "VolumeMAO_M12": mao.Range("M12").Value,
        "VolumeMAO_M19": mao.Range("M19").Value,
    }


def replace_component_code(component, code: str) -> None:
    module = component.CodeModule
    total_lines = module.CountOfLines
    if total_lines:
        module.DeleteLines(1, total_lines)
    module.AddFromString(code)


def ensure_component(vbproject, name: str, component_type: int):
    try:
        component = vbproject.VBComponents(name)
        if component.Type != component_type:
            vbproject.VBComponents.Remove(component)
            component = None
    except Exception:
        component = None

    if component is None:
        component = vbproject.VBComponents.Add(component_type)
        component.Name = name

    return component


def install_vba_cleanup(wb) -> None:
    project = wb.VBProject
    std_module = ensure_component(project, STD_MODULE_NAME, VBEXT_CT_STDMODULE)
    class_module = ensure_component(project, CLASS_MODULE_NAME, VBEXT_CT_CLASSMODULE)
    replace_component_code(std_module, STD_MODULE_CODE)
    replace_component_code(class_module, CLASS_MODULE_CODE)
    replace_component_code(project.VBComponents("ThisWorkbook"), THISWORKBOOK_CODE)


def kill_excel_process(pid: int | None) -> None:
    if not pid:
        return
    subprocess.run(
        ["taskkill", "/PID", str(pid), "/F"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=False,
    )


def main() -> int:
    if len(sys.argv) != 2:
        print("Usage: python scripts/remove_brado_embarcador_rows.py <workbook_path>", file=sys.stderr)
        return 2

    workbook_path = Path(sys.argv[1]).resolve()
    if not workbook_path.exists():
        print(f"Workbook not found: {workbook_path}", file=sys.stderr)
        return 2

    backup_dir = Path(__file__).resolve().parents[1] / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = backup_dir / f"{workbook_path.stem}_backup_{timestamp}{workbook_path.suffix}"
    shutil.copy2(workbook_path, backup_path)

    temp_dir = Path(tempfile.mkdtemp(prefix="codex_excel_"))
    temp_path = temp_dir / workbook_path.name

    pythoncom.CoInitialize()
    excel = None
    wb = None
    excel_pid: int | None = None
    result: dict[str, object] = {"backup": str(backup_path)}

    try:
        excel = win32.DispatchEx("Excel.Application")
        try:
            _, excel_pid = win32process.GetWindowThreadProcessId(excel.Hwnd)
        except Exception:
            excel_pid = None
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False

        wb = open_workbook(excel, workbook_path)
        install_vba_cleanup(wb)

        try:
            excel.Calculation = XL_CALC_MANUAL
        except Exception:
            pass

        cleanup_results = []
        for sheet_name in ("ROE_wk", "ROE_wk_monthly"):
            try:
                ws = wb.Worksheets(sheet_name)
            except Exception:
                continue
            cleanup_results.append(delete_brado_rows_from_sheet(ws))

        result["cleanup_results"] = cleanup_results
        result["removed_rows"] = sum(item["removed_rows"] for item in cleanup_results)

        try:
            excel.Calculation = XL_CALC_AUTOMATIC
        except Exception:
            pass

        excel.CalculateFullRebuild()
        result.update(workbook_metrics(wb))

        wb.SaveCopyAs(str(temp_path))
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    if excel_pid:
        for _ in range(5):
            try:
                os.kill(excel_pid, 0)
                time.sleep(1)
            except OSError:
                break
        else:
            kill_excel_process(excel_pid)

    for _ in range(15):
        try:
            os.replace(temp_path, workbook_path)
            break
        except PermissionError:
            time.sleep(2)
    else:
        raise PermissionError(f"Nao foi possivel substituir o arquivo original: {workbook_path}")

    pythoncom.CoInitialize()
    verify_excel = None
    verify_wb = None
    try:
        verify_excel = win32.DispatchEx("Excel.Application")
        verify_excel.Visible = False
        verify_excel.DisplayAlerts = False
        verify_wb = open_workbook(verify_excel, workbook_path)
        remaining_counts = {}
        for sheet_name in ("ROE_wk", "ROE_wk_monthly"):
            try:
                verify_ws = verify_wb.Worksheets(sheet_name)
            except Exception:
                continue
            headers = find_header_map(verify_ws)
            if "Embarcador" in headers:
                remaining_counts[sheet_name] = len(find_brado_rows(verify_ws, headers["Embarcador"]))
        result["remaining_brado_rows"] = remaining_counts
        result["verified_metrics"] = workbook_metrics(verify_wb)
    finally:
        if verify_wb is not None:
            try:
                verify_wb.Close(SaveChanges=False)
            except Exception:
                pass
        if verify_excel is not None:
            try:
                verify_excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
        shutil.rmtree(temp_dir, ignore_errors=True)

    print(json.dumps(result, ensure_ascii=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
