from __future__ import annotations

import os
import sys
from pathlib import Path

import pythoncom
import win32com.client

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from scripts.update_special_clients import WORKBOOK_PATH, backup_workbook, call_with_retry


XL_HIDDEN = 0
XL_ROW_FIELD = 1
XL_PAGE_FIELD = 3
XL_COUNT = -4112
XL_SUM = -4157

CUSTOMER_SHEET = "Top_Offenders_Customers"
SOURCE_PIVOT_NAME = "PivotTable1"
LEGACY_PIVOT_NAME = "PivotTable2"
HELPER_PIVOT_NAME = "PivotTable2_KPI_Helper"
HELPER_DESTINATION = "AF20"
LEGACY_DESTINATION = "AK20"
VISIBLE_CLEAR_RANGE = "T2:W2000"
VISIBLE_DATA_FIRST_ROW = 6
VISIBLE_MAX_ROWS = 1800
TARGET_WORKBOOK_PATH = Path(os.environ.get("DSU_WORKBOOK_PATH", str(WORKBOOK_PATH)))
LOG_PATH = ROOT_DIR / "analysis" / "align_top_offenders_customer_summary.log"


def log(message: str) -> None:
    print(f"[align-top-offenders] {message}", flush=True)
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    with LOG_PATH.open("a", encoding="utf-8") as handle:
        handle.write(f"{message}\n")


def pivot_table_by_name(worksheet, pivot_name: str):
    try:
        return worksheet.PivotTables(pivot_name)
    except Exception:
        return None


def normalize_week_value(value) -> str:
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def load_slicer_selection(workbook, cache_name: str) -> list[str]:
    try:
        cache = workbook.SlicerCaches.Item(cache_name)
    except Exception:
        return []

    selected = []
    try:
        items = cache.SlicerItems
        for index in range(1, items.Count + 1):
            item = items.Item(index)
            if getattr(item, "Selected", False):
                selected.append(str(item.Name))
    except Exception:
        return []
    return selected


def move_legacy_pivot_offscreen(worksheet) -> None:
    log("Checking legacy pivot position")
    legacy_pivot = pivot_table_by_name(worksheet, LEGACY_PIVOT_NAME)
    if legacy_pivot is None:
        log("Legacy pivot not found")
        return

    destination = worksheet.Range(LEGACY_DESTINATION)
    top_left = legacy_pivot.TableRange2.Cells(1, 1)
    if top_left.Row == destination.Row and top_left.Column == destination.Column:
        log("Legacy pivot already offscreen")
        return

    log("Moving legacy pivot offscreen")
    call_with_retry(legacy_pivot.TableRange2.Cut, destination)


def connect_slicer(workbook, cache_name: str, pivot_table) -> None:
    try:
        cache = workbook.SlicerCaches.Item(cache_name)
    except Exception:
        return
    try:
        cache.PivotTables.AddPivotTable(pivot_table)
    except Exception:
        pass


def ensure_row_field(pivot_table, field_name: str, position: int) -> None:
    log(f"Ensuring row field {field_name}")
    field = pivot_table.PivotFields(field_name)
    if field.Orientation != XL_ROW_FIELD:
        field.Orientation = XL_ROW_FIELD
    field.Position = position
    field.Subtotals = (False,) * 12


def hide_field(pivot_table, field_name: str) -> None:
    log(f"Hiding field {field_name}")
    field = pivot_table.PivotFields(field_name)
    if field.Orientation != XL_HIDDEN:
        field.Orientation = XL_HIDDEN


def ensure_page_field(pivot_table, field_name: str, page_value: str) -> None:
    log(f"Ensuring page field {field_name}={page_value}")
    field = pivot_table.PivotFields(field_name)
    if field.Orientation != XL_PAGE_FIELD:
        field.Orientation = XL_PAGE_FIELD
    field.Position = 1
    try:
        field.ClearAllFilters()
    except Exception:
        pass
    try:
        field.EnableMultiplePageItems = False
    except Exception:
        pass
    field.CurrentPage = page_value


def ensure_multi_page_field(pivot_table, field_name: str, visible_names: set[str]) -> None:
    if not visible_names:
        return
    log(f"Ensuring multi page field {field_name} with {len(visible_names)} visible items")
    field = pivot_table.PivotFields(field_name)
    if field.Orientation != XL_PAGE_FIELD:
        field.Orientation = XL_PAGE_FIELD
    field.Position = 1
    try:
        field.ClearAllFilters()
    except Exception:
        pass
    field.EnableMultiplePageItems = True
    items = field.PivotItems()
    for index in range(1, items.Count + 1):
        item = items.Item(index)
        if item.Name not in visible_names:
            try:
                item.Visible = False
            except Exception as exc:
                log(f"Could not hide item {field_name}::{item.Name!r}: {exc}")
                raise


def ensure_data_field(pivot_table, source_name: str, caption: str, function: int):
    log(f"Ensuring data field {caption}")
    data_fields = pivot_table.DataFields
    for index in range(1, data_fields.Count + 1):
        field = data_fields.Item(index)
        if getattr(field, "SourceName", "") == source_name:
            field.Function = function
            field.Caption = caption
            return field
    field = pivot_table.AddDataField(pivot_table.PivotFields(source_name))
    field.Function = function
    field.Caption = caption
    return field


def get_or_create_helper_pivot(worksheet):
    helper_pivot = pivot_table_by_name(worksheet, HELPER_PIVOT_NAME)
    if helper_pivot is not None:
        log("Reusing existing helper pivot")
        return helper_pivot

    log("Creating helper pivot")
    source_pivot = worksheet.PivotTables(SOURCE_PIVOT_NAME)
    return source_pivot.PivotCache().CreatePivotTable(
        worksheet.Range(HELPER_DESTINATION),
        HELPER_PIVOT_NAME,
    )


def find_helper_header_row(worksheet, pivot_table) -> int:
    log("Locating helper header row")
    top_row = pivot_table.TableRange2.Row
    row_count = pivot_table.TableRange2.Rows.Count
    for row in range(top_row, top_row + row_count):
        if worksheet.Cells(row, 32).Value == "Rótulos de Linha":
            return row
    raise RuntimeError("Could not find helper header row.")


def write_visible_summary(worksheet, header_row: int, active_week: str) -> None:
    log(f"Writing visible summary from helper header row {header_row}")
    end_row = VISIBLE_DATA_FIRST_ROW + VISIBLE_MAX_ROWS - 1
    call_with_retry(worksheet.Range(VISIBLE_CLEAR_RANGE).ClearContents)

    worksheet.Range("T2").Value = "Volume"
    worksheet.Range("U2").Value = "Ok"
    worksheet.Range("T3").Value = "OTO Out"
    worksheet.Range("U3").Value = "N"
    worksheet.Range("T4").Value = "Weeknum"
    worksheet.Range("U4").Value = active_week

    row_offset = header_row - VISIBLE_DATA_FIRST_ROW
    call_with_retry(
        setattr,
        worksheet.Range(f"T{VISIBLE_DATA_FIRST_ROW}:U{end_row}"),
        "FormulaR1C1",
        f"=R[{row_offset}]C[12]",
    )
    call_with_retry(
        setattr,
        worksheet.Range(f"W{VISIBLE_DATA_FIRST_ROW}:W{end_row}"),
        "FormulaR1C1",
        f"=R[{row_offset}]C[11]",
    )

    worksheet.Range("V6").Value = "OTO WK"
    call_with_retry(
        setattr,
        worksheet.Range(f"V7:V{end_row}"),
        "FormulaR1C1",
        '=IF(OR(RC[-1]="",RC[-1]=0,RC[1]=""),"",1-RC[1]/RC[-1])',
    )

    worksheet.Range("T6:W6").Font.Bold = True
    worksheet.Range(f"U7:U{end_row}").NumberFormat = "0"
    worksheet.Range(f"V7:V{end_row}").NumberFormat = "0%"
    worksheet.Range(f"W7:W{end_row}").NumberFormat = "0"


def configure_helper_pivot(workbook, worksheet) -> tuple[str, int]:
    log("Loading active week and slicer selections")
    active_week = normalize_week_value(workbook.Worksheets("Week_Overview").Range("AG1").Value)
    port_selection = load_slicer_selection(workbook, "Slicer_Porto1")

    move_legacy_pivot_offscreen(worksheet)
    helper_pivot = get_or_create_helper_pivot(worksheet)

    hide_field(helper_pivot, "Provedor")
    hide_field(helper_pivot, "Justificativa")
    ensure_row_field(helper_pivot, "Cliente Proposta", 1)

    ensure_page_field(helper_pivot, "Volume", "Ok")
    ensure_page_field(helper_pivot, "OTO Out", "N")
    ensure_page_field(helper_pivot, "Weeknum", active_week)

    if len(port_selection) == 1:
        ensure_page_field(helper_pivot, "Porto", port_selection[0])
    elif port_selection:
        ensure_multi_page_field(helper_pivot, "Porto", set(port_selection))

    ensure_multi_page_field(helper_pivot, "OTD ajustado", {"Atrasado", "No Prazo"})
    ensure_multi_page_field(helper_pivot, "Especiais", {""})

    ensure_data_field(helper_pivot, "Nº OS", "Count of OS", XL_COUNT)
    ensure_data_field(helper_pivot, "Atrasado?", "Sum of Atrasado?", XL_SUM)

    connect_slicer(workbook, "Slicer_Porto1", helper_pivot)
    connect_slicer(workbook, "NativeTimeline_Data_Prog.1", helper_pivot)

    log("Refreshing helper pivot")
    call_with_retry(helper_pivot.RefreshTable)
    worksheet.Columns("AF:AI").Hidden = True
    worksheet.Columns("AK:AN").Hidden = True
    return active_week, helper_pivot.TableRange2.Row


def inspect_visible_block(worksheet) -> list[list[object]]:
    rows = []
    for row in range(2, 15):
        rows.append([worksheet.Cells(row, col).Value for col in range(20, 24)])
    return rows


def open_workbook(excel):
    return call_with_retry(
        excel.Workbooks.Open,
        str(TARGET_WORKBOOK_PATH),
        UpdateLinks=0,
        ReadOnly=False,
        IgnoreReadOnlyRecommended=True,
        AddToMru=False,
        Notify=False,
    )


def main() -> int:
    if not TARGET_WORKBOOK_PATH.exists():
        log(f"Workbook not found: {TARGET_WORKBOOK_PATH}")
        return 1

    backup_path = backup_workbook(TARGET_WORKBOOK_PATH)
    log(f"Backup created at: {backup_path}")

    pythoncom.CoInitialize()
    excel = None
    workbook = None

    try:
        log("Starting Excel COM")
        excel = win32com.client.DispatchEx("Excel.Application")
        call_with_retry(setattr, excel, "Visible", False)
        call_with_retry(setattr, excel, "DisplayAlerts", False)
        call_with_retry(setattr, excel, "EnableEvents", False)
        call_with_retry(setattr, excel, "ScreenUpdating", False)

        log("Opening workbook - pass 1")
        workbook = open_workbook(excel)

        log("Workbook opened - pass 1")
        worksheet = workbook.Worksheets(CUSTOMER_SHEET)
        active_week, helper_top_row = configure_helper_pivot(workbook, worksheet)

        log("Recalculating and saving workbook - pass 1")
        call_with_retry(setattr, excel, "Calculation", -4105)
        call_with_retry(excel.CalculateFullRebuild)
        call_with_retry(workbook.Save)
        call_with_retry(workbook.Close, SaveChanges=False)
        workbook = None

        log("Reopening workbook - pass 2")
        workbook = open_workbook(excel)
        worksheet = workbook.Worksheets(CUSTOMER_SHEET)
        helper_pivot = pivot_table_by_name(worksheet, HELPER_PIVOT_NAME)
        if helper_pivot is None:
            raise RuntimeError("Helper pivot was not found on pass 2.")
        log("Refreshing helper pivot - pass 2")
        call_with_retry(helper_pivot.RefreshTable)
        header_row = find_helper_header_row(worksheet, helper_pivot)
        write_visible_summary(worksheet, header_row, active_week)
        log("Recalculating and saving workbook - pass 2")
        call_with_retry(setattr, excel, "Calculation", -4105)
        call_with_retry(excel.CalculateFullRebuild)
        call_with_retry(workbook.Save)

        visible_block = inspect_visible_block(worksheet)
        log(f"Applied active_week={active_week}, helper_top_row={helper_top_row}, helper_header_row={header_row}")
        log(f"Visible block preview: {visible_block}")
        return 0
    except Exception as exc:  # noqa: BLE001
        log(f"Failed to align customer summary pivot: {exc}")
        return 1
    finally:
        if workbook is not None:
            try:
                call_with_retry(workbook.Close, SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                call_with_retry(excel.Quit)
            except Exception:
                pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
