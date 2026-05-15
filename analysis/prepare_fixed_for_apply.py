from __future__ import annotations

import json
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

import pythoncom
import pywintypes
import win32com.client

ROOT = Path(r"C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX")
OFFICIAL = Path(r"C:\Users\VNO024\OneDrive - Maersk Group\Inland Execution Brazil - OPC - OPC (also BPS)\2026 - Info\DSU\Base_DSU2026 - TbM - WK19.xlsm")
FIXED = ROOT / "analysis" / "Base_DSU2026_WK19_rules_hypercare_fix_20260508_173020.xlsm"
REPORT = ROOT / "analysis" / "Base_DSU2026_WK19_prepare_for_apply_report.json"

XL_CALC_AUTOMATIC = -4105
XL_CELL_TYPE_LAST_CELL = 11
XL_MISSING_ITEMS_NONE = 0

def log(msg: str) -> None:
    print(f"[prepare-apply] {msg}", flush=True)

def call_with_retry(func, *args, retries: int = 30, delay: float = 0.4, **kwargs):
    last_exc = None
    for _ in range(retries):
        try:
            return func(*args, **kwargs)
        except pywintypes.com_error as exc:
            last_exc = exc
            time.sleep(delay)
            pythoncom.PumpWaitingMessages()
    raise last_exc

def safe_get(fn, default=None):
    try:
        return fn()
    except Exception:
        return default

def get_last_cell(ws):
    try:
        cell = call_with_retry(ws.Cells.SpecialCells, XL_CELL_TYPE_LAST_CELL)
        return int(cell.Row), int(cell.Column)
    except Exception:
        return 1, 1

def set_if_supported(obj, prop: str, value, stats: dict, key: str):
    try:
        setattr(obj, prop, value)
        stats[key] = stats.get(key, 0) + 1
        return True
    except Exception as exc:
        stats.setdefault(f"{key}_failed", 0)
        stats[f"{key}_failed"] += 1
        return False

def refresh_and_configure_fixed() -> dict:
    stats: dict = {
        "workbook": str(FIXED),
        "official_reference": str(OFFICIAL),
        "started_at": datetime.now().isoformat(timespec="seconds"),
        "connections": 0,
        "connection_refresh_attempts": 0,
        "connection_refresh_failures": [],
        "pivot_caches": 0,
        "pivot_cache_refresh_attempts": 0,
        "pivot_cache_refresh_failures": [],
        "pivot_tables": 0,
        "pivot_table_refresh_attempts": 0,
        "pivot_table_refresh_failures": [],
        "layout_sheets_processed": 0,
        "layout_columns_restored": 0,
        "layout_shapes_restored": 0,
        "layout_failures": [],
        "refresh_on_open_flags": 0,
        "refresh_on_open_flags_failed": 0,
        "preserve_formatting_flags": 0,
        "preserve_formatting_flags_failed": 0,
        "missing_items_limit_flags": 0,
        "missing_items_limit_flags_failed": 0,
    }

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    excel.AskToUpdateLinks = False

    official_wb = None
    fixed_wb = None
    try:
        log("Abrindo oficial como referencia visual...")
        official_wb = call_with_retry(
            excel.Workbooks.Open,
            str(OFFICIAL),
            UpdateLinks=False,
            ReadOnly=True,
        )
        log("Abrindo copia corrigida...")
        fixed_wb = call_with_retry(
            excel.Workbooks.Open,
            str(FIXED),
            UpdateLinks=False,
            ReadOnly=False,
        )

        call_with_retry(setattr, excel, "Calculation", XL_CALC_AUTOMATIC)
        set_if_supported(fixed_wb, "ForceFullCalculation", True, stats, "force_full_calculation_flags")
        set_if_supported(fixed_wb, "RefreshAllOnOpen", True, stats, "refresh_all_on_open_flags")

        # Configure workbook connections to refresh on open where supported.
        try:
            conn_count = int(fixed_wb.Connections.Count)
        except Exception:
            conn_count = 0
        stats["connections"] = conn_count
        for idx in range(1, conn_count + 1):
            try:
                conn = fixed_wb.Connections.Item(idx)
                for attr in ("OLEDBConnection", "ODBCConnection"):
                    obj = safe_get(lambda attr=attr: getattr(conn, attr))
                    if obj is not None:
                        set_if_supported(obj, "RefreshOnFileOpen", True, stats, "refresh_on_open_flags")
                        set_if_supported(obj, "BackgroundQuery", False, stats, "background_query_false_flags")
                try:
                    stats["connection_refresh_attempts"] += 1
                    conn.Refresh()
                except Exception as exc:
                    stats["connection_refresh_failures"].append({"index": idx, "error": str(exc)})
            except Exception as exc:
                stats["connection_refresh_failures"].append({"index": idx, "error": str(exc)})

        # Configure/refresh pivot caches.
        try:
            pc_count = int(fixed_wb.PivotCaches().Count)
        except Exception:
            pc_count = 0
        stats["pivot_caches"] = pc_count
        for idx in range(1, pc_count + 1):
            try:
                pc = fixed_wb.PivotCaches().Item(idx)
                set_if_supported(pc, "RefreshOnFileOpen", True, stats, "refresh_on_open_flags")
                set_if_supported(pc, "EnableRefresh", True, stats, "enable_refresh_flags")
                set_if_supported(pc, "MissingItemsLimit", XL_MISSING_ITEMS_NONE, stats, "missing_items_limit_flags")
                try:
                    stats["pivot_cache_refresh_attempts"] += 1
                    pc.Refresh()
                except Exception as exc:
                    stats["pivot_cache_refresh_failures"].append({"index": idx, "error": str(exc)})
            except Exception as exc:
                stats["pivot_cache_refresh_failures"].append({"index": idx, "error": str(exc)})

        # Refresh all and wait for async queries.
        log("Executando RefreshAll e recalculo completo...")
        try:
            fixed_wb.RefreshAll()
            stats["refresh_all_called"] = True
        except Exception as exc:
            stats["refresh_all_called"] = False
            stats["refresh_all_error"] = str(exc)
        try:
            excel.CalculateUntilAsyncQueriesDone()
            stats["calculate_until_async_queries_done"] = True
        except Exception as exc:
            stats["calculate_until_async_queries_done"] = False
            stats["calculate_until_async_queries_done_error"] = str(exc)
        try:
            excel.CalculateFullRebuild()
            stats["calculate_full_rebuild"] = True
        except Exception as exc:
            stats["calculate_full_rebuild"] = False
            stats["calculate_full_rebuild_error"] = str(exc)

        # Configure/refresh each pivot table and preserve visual formatting.
        for sidx in range(1, int(fixed_wb.Worksheets.Count) + 1):
            ws = fixed_wb.Worksheets.Item(sidx)
            ws_name = str(ws.Name)
            pt_count = safe_get(lambda: int(ws.PivotTables().Count), 0) or 0
            for pidx in range(1, pt_count + 1):
                try:
                    pt = ws.PivotTables().Item(pidx)
                    stats["pivot_tables"] += 1
                    set_if_supported(pt, "PreserveFormatting", True, stats, "preserve_formatting_flags")
                    set_if_supported(pt, "HasAutoFormat", False, stats, "has_autoformat_false_flags")
                    try:
                        stats["pivot_table_refresh_attempts"] += 1
                        pt.RefreshTable()
                    except Exception as exc:
                        stats["pivot_table_refresh_failures"].append({"sheet": ws_name, "index": pidx, "name": safe_get(lambda: str(pt.Name), "?"), "error": str(exc)})
                except Exception as exc:
                    stats["pivot_table_refresh_failures"].append({"sheet": ws_name, "index": pidx, "error": str(exc)})

        try:
            excel.CalculateFullRebuild()
        except Exception:
            pass

        # Restore visual column widths and shape positions from official.
        log("Restaurando larguras e posicoes visuais a partir do oficial...")
        fixed_names = {str(fixed_wb.Worksheets.Item(i).Name) for i in range(1, int(fixed_wb.Worksheets.Count) + 1)}
        for sidx in range(1, int(official_wb.Worksheets.Count) + 1):
            ows = official_wb.Worksheets.Item(sidx)
            name = str(ows.Name)
            if name not in fixed_names:
                stats["layout_failures"].append({"sheet": name, "error": "missing in fixed"})
                continue
            fws = fixed_wb.Worksheets(name)
            try:
                o_last_row, o_last_col = get_last_cell(ows)
                f_last_row, f_last_col = get_last_cell(fws)
                max_col = max(o_last_col, f_last_col, 80)
                max_col = min(max_col, 220)  # enough for this workbook and avoids Excel full-column loops
                for col in range(1, max_col + 1):
                    try:
                        ocol = ows.Columns(col)
                        fcol = fws.Columns(col)
                        fcol.ColumnWidth = ocol.ColumnWidth
                        fcol.Hidden = ocol.Hidden
                        try:
                            fcol.OutlineLevel = ocol.OutlineLevel
                        except Exception:
                            pass
                        stats["layout_columns_restored"] += 1
                    except Exception as exc:
                        stats["layout_failures"].append({"sheet": name, "col": col, "error": str(exc)})
                # Restore visible shape/chart positions by index when counts match or partially by min count.
                oshapes = ows.Shapes
                fshapes = fws.Shapes
                oc = int(oshapes.Count)
                fc = int(fshapes.Count)
                for shidx in range(1, min(oc, fc) + 1):
                    try:
                        osh = oshapes.Item(shidx)
                        fsh = fshapes.Item(shidx)
                        fsh.Left = osh.Left
                        fsh.Top = osh.Top
                        fsh.Width = osh.Width
                        fsh.Height = osh.Height
                        try:
                            fsh.Placement = osh.Placement
                        except Exception:
                            pass
                        stats["layout_shapes_restored"] += 1
                    except Exception as exc:
                        stats["layout_failures"].append({"sheet": name, "shape_index": shidx, "error": str(exc)})
                stats["layout_sheets_processed"] += 1
            except Exception as exc:
                stats["layout_failures"].append({"sheet": name, "error": str(exc)})

        log("Salvando copia corrigida preparada...")
        call_with_retry(fixed_wb.Save)
        stats["saved_after_prepare"] = True
    finally:
        if fixed_wb is not None:
            try:
                fixed_wb.Close(SaveChanges=False)
            except Exception:
                pass
        if official_wb is not None:
            try:
                official_wb.Close(SaveChanges=False)
            except Exception:
                pass
        try:
            excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()

    stats["finished_at"] = datetime.now().isoformat(timespec="seconds")
    return stats

def replace_vba_project_from_official(stats: dict) -> None:
    log("Preservando vbaProject.bin original do oficial na copia corrigida...")
    with ZipFile(OFFICIAL, "r") as z:
        official_vba = z.read("xl/vbaProject.bin")
    tmp = FIXED.with_suffix(".vba_tmp.xlsm")
    if tmp.exists():
        tmp.unlink()
    with ZipFile(FIXED, "r") as zin, ZipFile(tmp, "w", compression=ZIP_DEFLATED, allowZip64=True) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/vbaProject.bin":
                data = official_vba
            zout.writestr(item, data)
    shutil.move(str(tmp), str(FIXED))
    with ZipFile(FIXED, "r") as zfix, ZipFile(OFFICIAL, "r") as zoff:
        fixed_crc = zfix.getinfo("xl/vbaProject.bin").CRC
        official_crc = zoff.getinfo("xl/vbaProject.bin").CRC
        stats["vba_project_crc_matches_official"] = fixed_crc == official_crc
        stats["vba_project_crc_fixed"] = fixed_crc
        stats["vba_project_crc_official"] = official_crc


def main() -> int:
    if not OFFICIAL.exists():
        raise FileNotFoundError(OFFICIAL)
    if not FIXED.exists():
        raise FileNotFoundError(FIXED)
    # Safety copy of the fixed candidate before final prepare.
    pre = FIXED.with_name(FIXED.stem + "_before_prepare_for_apply_" + datetime.now().strftime("%Y%m%d_%H%M%S") + FIXED.suffix)
    shutil.copy2(FIXED, pre)
    stats = {"pre_prepare_candidate_backup": str(pre)}
    stats.update(refresh_and_configure_fixed())
    replace_vba_project_from_official(stats)
    REPORT.write_text(json.dumps(stats, ensure_ascii=False, indent=2), encoding="utf-8")
    log(f"Relatorio: {REPORT}")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
