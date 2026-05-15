from __future__ import annotations
import json, re, sys, time, traceback, hashlib
from pathlib import Path
from collections import Counter, defaultdict

import pythoncom
import win32com.client as win32

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

WORKBOOK = Path(r"C:\Users\VNO024\OneDrive - Maersk Group\2025\Py\DSU 2025\Base_DSU2026.xlsm")
OUT_JSON = Path(r"C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX\dsu_full_workbook_audit_result.json")

XL_CELL_TYPE_FORMULAS = -4123
XL_ERRORS = 16
XL_ROW_FIELD = 1
XL_COLUMN_FIELD = 2
XL_PAGE_FIELD = 3
XL_DATA_FIELD = 4

DASHBOARD_SHEETS = {
    "Volume_DS", "Volume_MAO", "Volume_Graph", "Week_Overview",
    "Top_Offenders_Customers", "Top_Offenders_Vendors", "Menu",
    "LastMonth_Overview", "Month2date"
}


def s(x):
    try:
        if x is None:
            return None
        return str(x)
    except Exception:
        return repr(x)


def try_get(obj, attr, default=None):
    try:
        return getattr(obj, attr)
    except Exception:
        return default


def count_range(rng):
    if rng is None:
        return 0
    try:
        return int(rng.CountLarge)
    except Exception:
        try:
            return int(rng.Count)
        except Exception:
            return 0


def sample_cells(rng, max_items=12):
    out = []
    if rng is None:
        return out
    try:
        for area in rng.Areas:
            rows = int(area.Rows.Count); cols = int(area.Columns.Count)
            for r in range(1, rows + 1):
                for c in range(1, cols + 1):
                    if len(out) >= max_items:
                        return out
                    cell = area.Cells(r, c)
                    out.append({
                        "address": s(cell.Address),
                        "text": s(cell.Text),
                        "value": s(cell.Value),
                        "formula": s(cell.Formula),
                    })
    except Exception as e:
        out.append({"sample_error": repr(e)})
    return out


def special_count_and_sample(used, cell_type, value_type=None, max_items=12):
    try:
        if value_type is None:
            rng = used.SpecialCells(cell_type)
        else:
            rng = used.SpecialCells(cell_type, value_type)
        return count_range(rng), sample_cells(rng, max_items)
    except Exception:
        return 0, []


def iter_matrix(value, rows, cols):
    if rows <= 0 or cols <= 0:
        return []
    if rows == 1 and cols == 1:
        return [[value]]
    if rows == 1:
        return [list(value)]
    if cols == 1:
        return [[r] for r in value]
    return [list(r) for r in value]


def matrix_summary(rng, max_cells_to_scan=250000):
    if rng is None:
        return {"exists": False, "rows": 0, "cols": 0, "blank_count": None, "error_count": None, "hash": None}
    rows = int(rng.Rows.Count); cols = int(rng.Columns.Count); total = rows * cols
    out = {"exists": True, "rows": rows, "cols": cols, "cell_count": total}
    if total > max_cells_to_scan:
        out.update({"skipped_scan": True, "reason": f"range too large ({total} cells)"})
        return out
    vals = iter_matrix(rng.Value, rows, cols)
    blanks = errors = nonblank = zeros = 0
    serial = []
    samples = []
    for ri, row in enumerate(vals, start=1):
        for ci, val in enumerate(row, start=1):
            sv = s(val)
            serial.append(sv)
            if val is None or val == "":
                blanks += 1
            else:
                nonblank += 1
                if isinstance(val, (int, float)) and val == 0:
                    zeros += 1
                if isinstance(val, str) and val.startswith("#"):
                    errors += 1
                    if len(samples) < 8:
                        samples.append({"rel_addr": f"R{ri}C{ci}", "value": sv})
    out.update({
        "blank_count": blanks,
        "nonblank_count": nonblank,
        "zero_count": zeros,
        "error_count": errors,
        "error_samples": samples,
        "hash": hashlib.md5(json.dumps(serial, ensure_ascii=False).encode("utf-8")).hexdigest(),
    })
    return out


def sheet_audit(ws):
    used = ws.UsedRange
    used_addr = s(used.Address)
    formula_count, _ = special_count_and_sample(used, XL_CELL_TYPE_FORMULAS, None, 0)
    err_count, err_samples = special_count_and_sample(used, XL_CELL_TYPE_FORMULAS, XL_ERRORS, 15)
    blank_formula_count = None
    blank_formula_samples = []
    # Only scan blank formula results on dashboard-sized sheets, not huge data tables.
    if ws.Name in DASHBOARD_SHEETS:
        try:
            rows = int(used.Rows.Count); cols = int(used.Columns.Count); total = rows * cols
            if total <= 120000:
                formulas = iter_matrix(used.Formula, rows, cols)
                values = iter_matrix(used.Value, rows, cols)
                blank_formula_count = 0
                for ri, row in enumerate(formulas, start=1):
                    for ci, formula in enumerate(row, start=1):
                        if isinstance(formula, str) and formula.startswith("="):
                            val = values[ri-1][ci-1]
                            if val is None or val == "":
                                blank_formula_count += 1
                                if len(blank_formula_samples) < 12:
                                    cell = used.Cells(ri, ci)
                                    blank_formula_samples.append({"address": s(cell.Address), "formula": formula[:180]})
        except Exception as e:
            blank_formula_samples = [{"scan_error": repr(e)}]
    try:
        chart_count = int(ws.ChartObjects().Count)
    except Exception:
        chart_count = 0
    try:
        pivot_count = int(ws.PivotTables().Count)
    except Exception:
        pivot_count = 0
    try:
        listobject_count = int(ws.ListObjects.Count)
    except Exception:
        listobject_count = 0
    return {
        "sheet": s(ws.Name),
        "used_range": used_addr,
        "used_rows": int(used.Rows.Count),
        "used_cols": int(used.Columns.Count),
        "formula_count": formula_count,
        "formula_errors_count": err_count,
        "formula_error_samples": err_samples,
        "blank_formula_results_count_dashboard_scan": blank_formula_count,
        "blank_formula_samples": blank_formula_samples,
        "pivot_count": pivot_count,
        "chart_count": chart_count,
        "listobject_count": listobject_count,
        "visible": s(try_get(ws, "Visible")),
    }


def pivot_filters(pt):
    filters = []
    try:
        fields = pt.PivotFields()
        for i in range(1, int(fields.Count) + 1):
            pf = fields.Item(i)
            try:
                orientation = int(pf.Orientation)
            except Exception:
                orientation = None
            # only record page fields, row/column/data fields, or fields with visible subset.
            item = {"name": s(pf.Name), "orientation": orientation}
            try:
                total = int(pf.PivotItems().Count)
                visible = []
                hidden_count = 0
                items = pf.PivotItems()
                for j in range(1, total + 1):
                    it = items.Item(j)
                    try:
                        is_vis = bool(it.Visible)
                    except Exception:
                        is_vis = True
                    if is_vis:
                        if len(visible) < 20:
                            visible.append(s(it.Name))
                    else:
                        hidden_count += 1
                item.update({"total_items": total, "visible_count": total - hidden_count, "hidden_count": hidden_count, "visible_items_first20": visible})
            except Exception:
                pass
            try:
                item["current_page"] = s(pf.CurrentPage.Name)
            except Exception:
                try:
                    item["current_page"] = s(pf.CurrentPage)
                except Exception:
                    pass
            if orientation in (XL_ROW_FIELD, XL_COLUMN_FIELD, XL_PAGE_FIELD, XL_DATA_FIELD) or item.get("hidden_count", 0) > 0:
                filters.append(item)
    except Exception as e:
        filters.append({"field_scan_error": repr(e)})
    return filters


def pivot_audit(ws, refresh=True):
    out = []
    try:
        pts = ws.PivotTables()
        for i in range(1, int(pts.Count) + 1):
            pt = pts.Item(i)
            pc = pt.PivotCache()
            try:
                body = pt.DataBodyRange
            except Exception:
                body = None
            before = matrix_summary(body)
            refresh_result = None
            refresh_error = None
            if refresh:
                try:
                    refresh_result = s(pt.RefreshTable())
                except Exception as e:
                    refresh_error = repr(e)
            try:
                body2 = pt.DataBodyRange
            except Exception:
                body2 = None
            after = matrix_summary(body2)
            changed = None
            if before.get("hash") and after.get("hash"):
                changed = before["hash"] != after["hash"]
            try:
                source = s(pc.SourceData)
            except Exception as e:
                source = f"<source error: {e}>"
            data_fields = []
            try:
                df = pt.DataFields
                if callable(df):
                    df = df()
                for j in range(1, int(df.Count) + 1):
                    f = df.Item(j)
                    data_fields.append({"name": s(f.Name), "source_name": s(try_get(f, "SourceName")), "function": s(try_get(f, "Function"))})
            except Exception:
                pass
            out.append({
                "sheet": s(ws.Name),
                "name": s(pt.Name),
                "range": s(pt.TableRange2.Address) if try_get(pt, "TableRange2") is not None else None,
                "data_body_range": s(body2.Address) if body2 is not None else None,
                "source_data": source,
                "cache_record_count": s(try_get(pc, "RecordCount")),
                "cache_refresh_date": s(try_get(pc, "RefreshDate")),
                "refresh_result": refresh_result,
                "refresh_error": refresh_error,
                "data_body_before": before,
                "data_body_after": after,
                "changed_after_refresh": changed,
                "data_fields": data_fields,
                "fields_filters": pivot_filters(pt),
            })
    except Exception as e:
        out.append({"sheet": s(ws.Name), "pivot_collection_error": repr(e)})
    return out


def flatten_values(values):
    if values is None:
        return []
    if isinstance(values, (str, int, float)):
        return [values]
    try:
        # COM tuple of tuples or tuple
        if isinstance(values, tuple):
            out = []
            for x in values:
                if isinstance(x, tuple):
                    out.extend(list(x))
                else:
                    out.append(x)
            return out
        return list(values)
    except Exception:
        return [values]


def series_values_summary(values):
    vals = flatten_values(values)
    errors = blanks = numbers = 0
    samples = []
    for i, v in enumerate(vals, start=1):
        sv = s(v)
        if v is None or v == "":
            blanks += 1
        elif isinstance(v, (int, float)):
            numbers += 1
        elif sv and sv.startswith("#"):
            errors += 1
            if len(samples) < 6:
                samples.append({"index": i, "value": sv})
    return {"count": len(vals), "blank_count": blanks, "number_count": numbers, "error_count": errors, "error_samples": samples}


def chart_audit(ws):
    out = []
    try:
        cos = ws.ChartObjects()
        for i in range(1, int(cos.Count) + 1):
            co = cos.Item(i)
            chart = co.Chart
            series = []
            try:
                sc = chart.SeriesCollection()
                sc_count = int(sc.Count)
                for j in range(1, sc_count + 1):
                    ser = sc.Item(j)
                    val_summary = {}
                    x_summary = {}
                    try:
                        val_summary = series_values_summary(ser.Values)
                    except Exception as e:
                        val_summary = {"values_error": repr(e)}
                    try:
                        x_summary = series_values_summary(ser.XValues)
                    except Exception as e:
                        x_summary = {"xvalues_error": repr(e)}
                    series.append({
                        "index": j,
                        "name": s(try_get(ser, "Name")),
                        "formula": s(try_get(ser, "Formula")),
                        "values_summary": val_summary,
                        "xvalues_summary": x_summary,
                    })
            except Exception as e:
                series.append({"series_collection_error": repr(e)})
            out.append({
                "sheet": s(ws.Name),
                "chart_object": s(co.Name),
                "chart_name": s(try_get(chart, "Name")),
                "chart_type": s(try_get(chart, "ChartType")),
                "series_count": len(series),
                "series": series,
            })
    except Exception as e:
        out.append({"sheet": s(ws.Name), "chart_collection_error": repr(e)})
    return out


def build_roe_records(wb):
    ws = wb.Worksheets("ROE_wk")
    lo = ws.ListObjects("ROE_wk")
    headers = [s(lo.ListColumns(i).Name) for i in range(1, int(lo.ListColumns.Count) + 1)]
    data = lo.DataBodyRange.Value
    rows = []
    for row in data:
        rec = {headers[i]: row[i] for i in range(len(headers))}
        rows.append(rec)
    return rows


def sv(v):
    if v is None or v == "":
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def count_recs(recs, *criteria):
    # criteria: (field, op, value); op eq/ne
    count = 0
    for r in recs:
        ok = True
        for field, op, value in criteria:
            rv = sv(r.get(field))
            vv = sv(value)
            if op == "eq" and rv != vv:
                ok = False; break
            if op == "ne" and rv == vv:
                ok = False; break
        if ok:
            count += 1
    return count


def oto_calc(recs, port=None, center=None):
    criteria_den = [("Volume", "eq", "Ok"), ("Weeknum", "eq", 20), ("OTD ajustado", "ne", "Sem Preenchimento"), ("OTO Out", "eq", "N")]
    criteria_num = [("Volume", "eq", "Ok"), ("Weeknum", "eq", 20), ("OTD ajustado", "eq", "Atrasado"), ("OTO Out", "eq", "N"), ("Especiais", "ne", "Especial")]
    if port is not None:
        criteria_den.append(("Porto", "eq", port)); criteria_num.append(("Porto", "eq", port))
    if center == "CAB":
        criteria_den.append(("Centro de Custo", "eq", "Aliança")); criteria_num.append(("Centro de Custo", "eq", "Aliança"))
    elif center == "DS":
        criteria_den.append(("Centro de Custo", "ne", "Aliança")); criteria_num.append(("Centro de Custo", "ne", "Aliança"))
    den = count_recs(recs, *criteria_den)
    num = count_recs(recs, *criteria_num)
    return {"denominator": den, "delayed_penalty": num, "oto": None if den == 0 else 1 - num / den}


def moves_count(recs, port=None, center=None):
    criteria = [("Volume", "eq", "Ok"), ("Weeknum", "eq", 20)]
    if port is not None:
        criteria.append(("Porto", "eq", port))
    if center == "CAB":
        criteria.append(("Centro de Custo", "eq", "Aliança"))
    elif center == "DS":
        criteria.append(("Centro de Custo", "ne", "Aliança"))
    return count_recs(recs, *criteria)


def week_overview_validation(wb, recs):
    ws = wb.Worksheets("Week_Overview")
    out = []
    rows = [2,3,6,7,8,11,12,13,16,17,18,19,20,23]
    for row in rows:
        label = s(ws.Range(f"A{row}").Value)
        port = s(ws.Range(f"AG{row}").Value)
        if row == 23:
            port = None
        calc = {
            "row": row,
            "label": label,
            "port_filter": port,
            "moves_cab_calc": moves_count(recs, port, "CAB"),
            "moves_ds_calc": moves_count(recs, port, "DS"),
            "moves_total_calc": moves_count(recs, port, None),
            "oto_cab_calc": oto_calc(recs, port, "CAB"),
            "oto_ds_calc": oto_calc(recs, port, "DS"),
            "oto_total_calc": oto_calc(recs, port, None),
            "sheet_values": {
                "D_moves_cab": ws.Range(f"D{row}").Value,
                "H_moves_ds": ws.Range(f"H{row}").Value,
                "K_moves_total": ws.Range(f"K{row}").Value,
                "N_oto_cab": ws.Range(f"N{row}").Value,
                "P_oto_ds": ws.Range(f"P{row}").Value,
                "Q_oto_total": ws.Range(f"Q{row}").Value,
            }
        }
        # differences, with rounding tolerance for rounded total rows
        diffs = []
        if calc["sheet_values"]["D_moves_cab"] not in (None, "") and abs(float(calc["sheet_values"]["D_moves_cab"]) - calc["moves_cab_calc"]) > 0.001:
            diffs.append("D moves CAB")
        if calc["sheet_values"]["H_moves_ds"] not in (None, "") and abs(float(calc["sheet_values"]["H_moves_ds"]) - calc["moves_ds_calc"]) > 0.001:
            diffs.append("H moves DS")
        if calc["sheet_values"]["K_moves_total"] not in (None, "") and abs(float(calc["sheet_values"]["K_moves_total"]) - calc["moves_total_calc"]) > 0.001:
            diffs.append("K moves total")
        for key, calc_key in [("N_oto_cab", "oto_cab_calc"), ("P_oto_ds", "oto_ds_calc"), ("Q_oto_total", "oto_total_calc")]:
            sheet_v = calc["sheet_values"][key]
            calc_v = calc[calc_key]["oto"]
            if calc_v is None:
                if sheet_v not in (None, ""):
                    diffs.append(key)
            else:
                # N23/P23 are rounded in sheet; allow 0.006
                tol = 0.006 if row == 23 and key in ("N_oto_cab", "P_oto_ds") else 0.000001
                if sheet_v in (None, "") or abs(float(sheet_v) - calc_v) > tol:
                    diffs.append(key)
        calc["differences"] = diffs
        out.append(calc)
    return out


def main():
    result = {"workbook": str(WORKBOOK), "started_at": time.strftime("%Y-%m-%d %H:%M:%S"), "issues": []}
    pythoncom.CoInitialize(); excel = None; wb = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        try: excel.AskToUpdateLinks = False
        except Exception: pass
        try: excel.AutomationSecurity = 3
        except Exception: pass
        result["excel_version"] = s(excel.Version)
        wb = excel.Workbooks.Open(str(WORKBOOK), UpdateLinks=0, ReadOnly=True, IgnoreReadOnlyRecommended=True, AddToMru=False)
        result["opened_read_only"] = bool(wb.ReadOnly)
        result["workbook_full_name"] = s(wb.FullName)
        result["sheets"] = []
        result["pivots"] = []
        result["charts"] = []
        for idx in range(1, int(wb.Worksheets.Count) + 1):
            ws = wb.Worksheets(idx)
            result["sheets"].append(sheet_audit(ws))
        # Refresh pivots in memory and inspect charts after calculation.
        for idx in range(1, int(wb.Worksheets.Count) + 1):
            ws = wb.Worksheets(idx)
            pivs = pivot_audit(ws, refresh=True)
            result["pivots"].extend(pivs)
        try:
            excel.CalculateFullRebuild()
        except Exception as e:
            result["calculate_error"] = repr(e)
        for idx in range(1, int(wb.Worksheets.Count) + 1):
            ws = wb.Worksheets(idx)
            try:
                if int(ws.ChartObjects().Count) > 0:
                    result["charts"].extend(chart_audit(ws))
            except Exception:
                pass
        # Chart sheets, if any.
        result["chart_sheets"] = []
        try:
            for i in range(1, int(wb.Charts.Count) + 1):
                ch = wb.Charts(i)
                result["chart_sheets"].append({"name": s(ch.Name), "chart_type": s(try_get(ch, "ChartType"))})
        except Exception as e:
            result["chart_sheet_error"] = repr(e)
        # Independent Week_Overview validation from ROE_wk listobject.
        recs = build_roe_records(wb)
        result["roe_summary"] = {
            "rows": len(recs),
            "week20_volume_ok": count_recs(recs, ("Volume", "eq", "Ok"), ("Weeknum", "eq", 20)),
            "week20_volume_ok_cab": count_recs(recs, ("Volume", "eq", "Ok"), ("Weeknum", "eq", 20), ("Centro de Custo", "eq", "Aliança")),
            "week20_volume_ok_ds": count_recs(recs, ("Volume", "eq", "Ok"), ("Weeknum", "eq", 20), ("Centro de Custo", "ne", "Aliança")),
            "week20_atrasado_penalty": count_recs(recs, ("Volume", "eq", "Ok"), ("Weeknum", "eq", 20), ("OTD ajustado", "eq", "Atrasado"), ("OTO Out", "eq", "N"), ("Especiais", "ne", "Especial")),
            "week20_sem_preenchimento": count_recs(recs, ("Volume", "eq", "Ok"), ("Weeknum", "eq", 20), ("OTD ajustado", "eq", "Sem Preenchimento")),
        }
        result["week_overview_validation"] = week_overview_validation(wb, recs)
        # Summaries/issues
        formula_error_sheets = [x for x in result["sheets"] if x.get("formula_errors_count", 0) > 0]
        for x in formula_error_sheets:
            result["issues"].append({"severity": "attention", "type": "formula_errors", "sheet": x["sheet"], "count": x["formula_errors_count"], "samples": x.get("formula_error_samples", [])[:5]})
        empty_pivots = [p for p in result["pivots"] if p.get("data_body_after", {}).get("exists") is False or p.get("refresh_error")]
        for p in empty_pivots:
            result["issues"].append({"severity": "warning", "type": "pivot_empty_or_refresh_error", "sheet": p.get("sheet"), "pivot": p.get("name"), "refresh_error": p.get("refresh_error")})
        stale_pivots = [p for p in result["pivots"] if p.get("changed_after_refresh") is True]
        for p in stale_pivots:
            result["issues"].append({"severity": "attention", "type": "pivot_changed_after_refresh", "sheet": p.get("sheet"), "pivot": p.get("name")})
        chart_issues = []
        for ch in result["charts"]:
            if ch.get("series_count", 0) == 0:
                chart_issues.append({"chart": ch.get("chart_object"), "sheet": ch.get("sheet"), "reason": "no series"})
            for ser in ch.get("series", []):
                vs = ser.get("values_summary", {})
                if vs.get("count", 0) == 0 or vs.get("error_count", 0) > 0 or (vs.get("blank_count", 0) == vs.get("count", -1) and vs.get("count", 0) > 0):
                    chart_issues.append({"chart": ch.get("chart_object"), "sheet": ch.get("sheet"), "series": ser.get("index"), "name": ser.get("name"), "values_summary": vs, "formula": ser.get("formula")})
        for ci in chart_issues:
            result["issues"].append({"severity": "attention", "type": "chart_series_issue", **ci})
        wk_diffs = [x for x in result["week_overview_validation"] if x.get("differences")]
        for d in wk_diffs:
            result["issues"].append({"severity": "warning", "type": "week_overview_mismatch", "row": d["row"], "label": d["label"], "differences": d["differences"]})
        result["finished_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
        OUT_JSON.write_text(json.dumps(result, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
        return result
    except Exception as e:
        result["fatal_error"] = repr(e)
        result["traceback"] = traceback.format_exc()
        OUT_JSON.write_text(json.dumps(result, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
        return result
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    res = main()
    print("WROTE", OUT_JSON)
    print(json.dumps({
        "fatal_error": res.get("fatal_error"),
        "opened_read_only": res.get("opened_read_only"),
        "sheet_count": len(res.get("sheets", [])),
        "pivot_count": len([p for p in res.get("pivots", []) if p.get("name")]),
        "chart_count": len(res.get("charts", [])),
        "issue_count": len(res.get("issues", [])),
        "roe_summary": res.get("roe_summary"),
        "issues_preview": res.get("issues", [])[:20],
    }, ensure_ascii=False, indent=2, default=str))
    if res.get("fatal_error"):
        raise SystemExit(1)
