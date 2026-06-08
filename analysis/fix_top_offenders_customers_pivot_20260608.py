from __future__ import annotations

import json
import math
import shutil
import time
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client as win32
import win32process

ROOT = Path(r"C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX")
OFFICIAL = Path(r"C:\Users\VNO024\OneDrive - Maersk Group\Inland Execution Brazil - OPC - OPC (also BPS)\2026 - Info\DSU\Base_DSU2026 - TbM - WK23.xlsm")
ANALYSIS = ROOT / "analysis"
BACKUPS = ROOT / "backups"
SHEET_NAME = "Top_Offenders_Customers"
PIVOT_NAME = "PivotTable1"
CUSTOMER_FIELD = "Cliente Proposta"
WEEK_OVERVIEW = "Week_Overview"
ROE_SHEET = "ROE_wk"

# Excel constants used directly to avoid depending on generated wrappers.
XL_DONE = 0
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3


def as_jsonable(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, (str, int, bool)):
        return value
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return str(value)
        return value
    if isinstance(value, (datetime,)):
        return value.isoformat()
    if isinstance(value, (list, tuple)):
        return [as_jsonable(v) for v in value]
    if isinstance(value, dict):
        return {str(k): as_jsonable(v) for k, v in value.items()}
    return str(value)


def normalize_range_value(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, tuple):
        return [[values]]
    if values and not isinstance(values[0], tuple):
        return [list(values)]
    return [list(row) for row in values]


def excel_start() -> tuple[Any, int | None]:
    pythoncom.CoInitialize()
    xl = win32.DispatchEx("Excel.Application")
    pid = None
    try:
        _, pid = win32process.GetWindowThreadProcessId(xl.Hwnd)
    except Exception:
        pid = None
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AskToUpdateLinks = False
    try:
        xl.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
    except Exception:
        pass
    return xl, pid


def excel_stop(xl: Any) -> None:
    try:
        xl.DisplayAlerts = False
    except Exception:
        pass
    try:
        xl.EnableEvents = True
    except Exception:
        pass
    try:
        xl.Quit()
    except Exception:
        pass
    try:
        pythoncom.CoUninitialize()
    except Exception:
        pass


def wait_calculation(xl: Any, timeout_s: int = 120) -> int | str:
    start = time.time()
    last_state: int | str = "unknown"
    while time.time() - start < timeout_s:
        try:
            last_state = xl.CalculationState
            if last_state == XL_DONE:
                return last_state
        except Exception as exc:
            last_state = f"error: {exc}"
            break
        time.sleep(0.5)
    return last_state


def open_workbook(xl: Any, path: Path, read_only: bool) -> Any:
    return xl.Workbooks.Open(
        Filename=str(path),
        UpdateLinks=0,
        ReadOnly=read_only,
        IgnoreReadOnlyRecommended=True,
        AddToMru=False,
    )


def range_values(rng: Any, max_rows: int | None = None) -> list[list[Any]]:
    values = normalize_range_value(rng.Value)
    if max_rows is not None:
        return values[:max_rows]
    return values


def numeric_values(row: list[Any]) -> list[float]:
    out = []
    for value in row:
        if isinstance(value, bool):
            continue
        if isinstance(value, (int, float)):
            out.append(float(value))
    return out


def range_address(rng: Any) -> str:
    try:
        return rng.Address(False, False)
    except Exception:
        try:
            return str(rng.Address)
        except Exception as exc:
            return f"address_error: {exc}"


def find_total_in_pivot(pt: Any) -> dict[str, Any]:
    table = pt.TableRange2
    values = range_values(table)
    candidates: list[dict[str, Any]] = []
    for row_idx, row in enumerate(values, start=1):
        texts = [str(v).strip() for v in row if v is not None and str(v).strip()]
        joined = " | ".join(texts).lower()
        nums = numeric_values(row)
        if nums and ("total geral" in joined or "grand total" in joined or joined.startswith("total") or " total" in joined):
            candidates.append(
                {
                    "relative_row": row_idx,
                    "label": texts[0] if texts else None,
                    "numbers": nums,
                    # In this pivot, the first numeric column is the OS/volume count.
                    # Later numeric columns are KPI metrics, not the volume total.
                    "total": nums[0],
                    "volume_total": nums[0],
                    "last_numeric_metric": nums[-1],
                }
            )
    fallback = None
    if values:
        last = values[-1]
        nums = numeric_values(last)
        texts = [str(v).strip() for v in last if v is not None and str(v).strip()]
        if nums:
            fallback = {
                "relative_row": len(values),
                "label": texts[0] if texts else None,
                "numbers": nums,
                "total": nums[0],
                "volume_total": nums[0],
                "last_numeric_metric": nums[-1],
                "fallback": True,
            }
    chosen = candidates[-1] if candidates else fallback
    return {
        "table_address": range_address(table),
        "rows": table.Rows.Count,
        "columns": table.Columns.Count,
        "total_candidates": candidates[-5:],
        "chosen_total": chosen,
    }


def pivot_field_counts(pf: Any, sample_limit: int = 20) -> dict[str, Any]:
    count = int(pf.PivotItems().Count)
    visible = 0
    hidden = 0
    visible_samples: list[str] = []
    hidden_samples: list[str] = []
    errors: list[str] = []
    for idx in range(1, count + 1):
        try:
            item = pf.PivotItems(idx)
            name = str(item.Name)
            is_visible = bool(item.Visible)
            if is_visible:
                visible += 1
                if len(visible_samples) < sample_limit:
                    visible_samples.append(name)
            else:
                hidden += 1
                if len(hidden_samples) < sample_limit:
                    hidden_samples.append(name)
        except Exception as exc:
            errors.append(f"item {idx}: {exc}")
            if len(errors) >= 10:
                break
    return {
        "items": count,
        "visible": visible,
        "hidden": hidden,
        "visible_samples": visible_samples,
        "hidden_samples": hidden_samples,
        "errors": errors,
    }


def page_field_snapshot(pt: Any) -> list[dict[str, Any]]:
    fields = []
    try:
        count = int(pt.PivotFields().Count)
    except Exception:
        return fields
    for idx in range(1, count + 1):
        try:
            pf = pt.PivotFields(idx)
            orientation = int(pf.Orientation)
            if orientation == 3:  # xlPageField
                current_page = None
                try:
                    current_page = str(pf.CurrentPage.Name)
                except Exception:
                    try:
                        current_page = str(pf.CurrentPage)
                    except Exception as exc:
                        current_page = f"error: {exc}"
                fields.append(
                    {
                        "name": str(pf.Name),
                        "current_page": current_page,
                        "enable_multiple_page_items": bool(getattr(pf, "EnableMultiplePageItems", False)),
                    }
                )
        except Exception as exc:
            fields.append({"index": idx, "error": str(exc)})
    return fields


def get_week_overview_expected(wb: Any) -> dict[str, Any]:
    ws = wb.Worksheets(WEEK_OVERVIEW)
    active_week = int(float(ws.Range("AI1").Value))
    row = active_week
    return {
        "active_week": active_week,
        "row_used": row,
        "cab": ws.Range(f"D{row}").Value,
        "ds": ws.Range(f"H{row}").Value,
        "total": ws.Range(f"K{row}").Value,
        "oto_out": ws.Range(f"N{row}").Value,
        "otd_ajustado": ws.Range(f"P{row}").Value,
        "performance": ws.Range(f"Q{row}").Value,
    }


def norm_header(value: Any) -> str:
    return str(value).strip().lower() if value is not None else ""


def as_int_or_str(value: Any) -> Any:
    if value is None or value == "":
        return value
    try:
        return int(float(value))
    except Exception:
        return str(value).strip()


def compute_roe_expected(wb: Any, active_week: int) -> dict[str, Any]:
    result: dict[str, Any] = {"sheet": ROE_SHEET, "status": "not_run"}
    try:
        ws = wb.Worksheets(ROE_SHEET)
        values = range_values(ws.UsedRange)
        header_row_idx = None
        headers: list[str] = []
        for idx, row in enumerate(values[:20]):
            normalized = [norm_header(v) for v in row]
            if "volume" in normalized and ("weeknum" in normalized or "week num" in normalized or "week" in normalized):
                header_row_idx = idx
                headers = normalized
                break
            if "volume" in normalized:
                header_row_idx = idx
                headers = normalized
                break
        if header_row_idx is None:
            result.update({"status": "header_not_found"})
            return result
        def col(*names: str) -> int | None:
            for name in names:
                if name.lower() in headers:
                    return headers.index(name.lower())
            return None
        idx_volume = col("volume")
        idx_week = col("weeknum", "week num", "week")
        idx_os = col("os", "ordem", "ordem de servico", "ordem de serviço")
        idx_mod = col("Modal", "Modo")
        if idx_volume is None:
            result.update({"status": "volume_column_not_found", "headers_sample": headers[:80]})
            return result
        rows_ok = 0
        unique_os: set[str] = set()
        cab = 0
        ds = 0
        for row in values[header_row_idx + 1:]:
            def get(i: int | None) -> Any:
                if i is None or i >= len(row):
                    return None
                return row[i]
            volume = str(get(idx_volume)).strip()
            if volume != "Ok":
                continue
            if idx_week is not None:
                wk = as_int_or_str(get(idx_week))
                if wk != active_week:
                    continue
            rows_ok += 1
            os_value = get(idx_os)
            if os_value is not None and str(os_value).strip():
                unique_os.add(str(os_value).strip())
            mod = str(get(idx_mod)).strip().upper() if idx_mod is not None and get(idx_mod) is not None else ""
            if mod == "CAB":
                cab += 1
            elif mod == "DS":
                ds += 1
        result.update(
            {
                "status": "ok",
                "header_row_1based": header_row_idx + 1,
                "rows_volume_ok": rows_ok,
                "unique_os_volume_ok": len(unique_os) if idx_os is not None else None,
                "cab_rows": cab if idx_mod is not None else None,
                "ds_rows": ds if idx_mod is not None else None,
                "used_rows": len(values),
                "used_cols": max((len(r) for r in values), default=0),
                "columns_found": {
                    "volume": idx_volume + 1 if idx_volume is not None else None,
                    "week": idx_week + 1 if idx_week is not None else None,
                    "os": idx_os + 1 if idx_os is not None else None,
                    "modal": idx_mod + 1 if idx_mod is not None else None,
                },
            }
        )
    except Exception as exc:
        result.update({"status": "error", "error": str(exc)})
    return result


def snapshot_workbook(path: Path, read_only: bool = True) -> dict[str, Any]:
    xl, pid = excel_start()
    wb = None
    try:
        wb = open_workbook(xl, path, read_only=read_only)
        week = get_week_overview_expected(wb)
        roe = compute_roe_expected(wb, int(week["active_week"]))
        ws = wb.Worksheets(SHEET_NAME)
        pt = ws.PivotTables(PIVOT_NAME)
        pf = pt.PivotFields(CUSTOMER_FIELD)
        return {
            "path": str(path),
            "excel_pid": pid,
            "read_only_requested": read_only,
            "workbook_read_only": bool(wb.ReadOnly),
            "week_overview": week,
            "roe_expected": roe,
            "pivot": {
                "sheet": SHEET_NAME,
                "name": PIVOT_NAME,
                "page_fields": page_field_snapshot(pt),
                "customer_field_counts": pivot_field_counts(pf),
                "total_snapshot": find_total_in_pivot(pt),
            },
        }
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        excel_stop(xl)


def set_all_customer_items_visible(pf: Any) -> dict[str, Any]:
    set_visible = 0
    errors: list[str] = []
    count = int(pf.PivotItems().Count)
    for idx in range(1, count + 1):
        try:
            item = pf.PivotItems(idx)
            if not bool(item.Visible):
                item.Visible = True
                set_visible += 1
        except Exception as exc:
            errors.append(f"{idx}: {exc}")
            if len(errors) >= 20:
                break
    return {"attempted_items": count, "set_visible": set_visible, "errors": errors}


def apply_fix_to_workbook(path: Path) -> dict[str, Any]:
    xl, pid = excel_start()
    wb = None
    result: dict[str, Any] = {"path": str(path), "excel_pid": pid, "steps": []}
    try:
        wb = open_workbook(xl, path, read_only=False)
        result["workbook_read_only"] = bool(wb.ReadOnly)
        if wb.ReadOnly:
            raise RuntimeError("Workbook abriu como somente leitura; abortando para não fingir que salvou.")
        week = get_week_overview_expected(wb)
        result["week_overview"] = week
        result["roe_expected"] = compute_roe_expected(wb, int(week["active_week"]))
        ws = wb.Worksheets(SHEET_NAME)
        pt = ws.PivotTables(PIVOT_NAME)
        pf = pt.PivotFields(CUSTOMER_FIELD)
        before_counts = pivot_field_counts(pf)
        before_total = find_total_in_pivot(pt)
        before_pages = page_field_snapshot(pt)
        result["before"] = {
            "page_fields": before_pages,
            "customer_field_counts": before_counts,
            "total_snapshot": before_total,
        }
        # Clear manual/label/value filters on the customer row field. This is the intended scoped fix.
        result["steps"].append("ClearAllFilters em Cliente Proposta na PivotTable1")
        try:
            pt.ManualUpdate = True
        except Exception:
            pass
        try:
            pf.ClearAllFilters()
        finally:
            try:
                pt.ManualUpdate = False
            except Exception:
                pass
        after_clear_counts = pivot_field_counts(pf)
        visibility_step = {"after_clear_counts": after_clear_counts}
        if after_clear_counts.get("hidden", 0):
            result["steps"].append("Forçar itens ocultos restantes para Visible=True")
            visibility_step["force_visible"] = set_all_customer_items_visible(pf)
        result["visibility_step"] = visibility_step
        # Refresh only the changed pivot table/cache, not all workbook queries.
        result["steps"].append("RefreshTable apenas da PivotTable1")
        try:
            pt.RefreshTable()
        except Exception as exc:
            result["refresh_warning"] = str(exc)
        try:
            xl.CalculateFullRebuild()
        except Exception:
            try:
                xl.CalculateFull()
            except Exception as exc:
                result["calc_warning"] = str(exc)
        result["calculation_state"] = wait_calculation(xl)
        after_counts = pivot_field_counts(pf)
        after_total = find_total_in_pivot(pt)
        after_pages = page_field_snapshot(pt)
        result["after"] = {
            "page_fields": after_pages,
            "customer_field_counts": after_counts,
            "total_snapshot": after_total,
        }
        chosen = after_total.get("chosen_total") or {}
        after_value = chosen.get("total")
        expected_total = week.get("ds")
        validations = {
            "customer_hidden_items_zero": after_counts.get("hidden") == 0,
            "customer_visible_equals_items": after_counts.get("visible") == after_counts.get("items"),
            "pivot_total": after_value,
            "expected_week_overview_ds": expected_total,
            "week_overview_total_all": week.get("total"),
            "pivot_total_matches_week_overview_ds": round(float(after_value or 0), 6) == round(float(expected_total or 0), 6),
        }
        roe = result.get("roe_expected", {})
        roe_unique = roe.get("unique_os_volume_ok") if isinstance(roe, dict) else None
        if roe_unique is not None:
            validations["expected_roe_unique_os"] = roe_unique
            validations["pivot_total_matches_roe_unique_os"] = round(float(after_value or 0), 6) == round(float(roe_unique or 0), 6)
        result["validations"] = validations
        if not validations["customer_hidden_items_zero"] or not validations["pivot_total_matches_week_overview_ds"]:
            raise RuntimeError(f"Validação falhou antes de salvar: {validations}")
        result["steps"].append("Salvar workbook")
        wb.Save()
        result["saved"] = True
        return result
    except Exception as exc:
        result["error"] = str(exc)
        result["traceback"] = traceback.format_exc()
        raise RuntimeError(json.dumps(as_jsonable(result), ensure_ascii=False, indent=2)) from exc
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        excel_stop(xl)


def main() -> None:
    ANALYSIS.mkdir(parents=True, exist_ok=True)
    BACKUPS.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    test_copy = ANALYSIS / f"Base_DSU2026 - TbM - WK23_top_offenders_customers_test_{stamp}.xlsm"
    backup = BACKUPS / f"Base_DSU2026 - TbM - WK23_before_top_offenders_customers_fix_{stamp}.xlsm"
    report_path = ANALYSIS / f"top_offenders_customers_fix_apply_{stamp}.json"
    md_path = ANALYSIS / f"top_offenders_customers_fix_apply_{stamp}.md"

    report: dict[str, Any] = {
        "timestamp": stamp,
        "official_path": str(OFFICIAL),
        "test_copy": str(test_copy),
        "backup_path": str(backup),
        "scope": {
            "sheet": SHEET_NAME,
            "pivot": PIVOT_NAME,
            "field_cleared": CUSTOMER_FIELD,
            "not_changed": ["PivotTable5", "Tabela dinâmica1", "PivotTable2_KPI_Helper", "PivotTable2", "Top_Offenders_Vendors", "Week_Overview", "Region errors"],
        },
    }

    if not OFFICIAL.exists():
        raise FileNotFoundError(str(OFFICIAL))

    report["baseline_official"] = snapshot_workbook(OFFICIAL, read_only=True)

    shutil.copy2(OFFICIAL, test_copy)
    report["test_apply"] = apply_fix_to_workbook(test_copy)
    report["test_reopen_validation"] = snapshot_workbook(test_copy, read_only=True)

    test_total = (((report["test_reopen_validation"].get("pivot") or {}).get("total_snapshot") or {}).get("chosen_total") or {}).get("total")
    expected_total = (report["test_reopen_validation"].get("week_overview") or {}).get("ds")
    if round(float(test_total or 0), 6) != round(float(expected_total or 0), 6):
        raise RuntimeError(f"Cópia de teste não fechou contra DS: pivot={test_total} esperado={expected_total}")

    shutil.copy2(OFFICIAL, backup)
    report["backup_created"] = {"path": str(backup), "size": backup.stat().st_size}
    report["official_apply"] = apply_fix_to_workbook(OFFICIAL)
    report["official_reopen_validation"] = snapshot_workbook(OFFICIAL, read_only=True)

    official_total = (((report["official_reopen_validation"].get("pivot") or {}).get("total_snapshot") or {}).get("chosen_total") or {}).get("total")
    official_expected = (report["official_reopen_validation"].get("week_overview") or {}).get("ds")
    official_counts = ((report["official_reopen_validation"].get("pivot") or {}).get("customer_field_counts") or {})
    report["final_validation"] = {
        "official_total": official_total,
        "expected_week_overview_ds": official_expected,
        "week_overview_total_all": (report["official_reopen_validation"].get("week_overview") or {}).get("total"),
        "matches_week_overview_ds": round(float(official_total or 0), 6) == round(float(official_expected or 0), 6),
        "customer_items": official_counts.get("items"),
        "customer_visible": official_counts.get("visible"),
        "customer_hidden": official_counts.get("hidden"),
        "hidden_zero": official_counts.get("hidden") == 0,
    }

    report_path.write_text(json.dumps(as_jsonable(report), ensure_ascii=False, indent=2), encoding="utf-8")

    def val(path_list: list[str], default: Any = None) -> Any:
        node: Any = report
        for key in path_list:
            if not isinstance(node, dict):
                return default
            node = node.get(key)
        return node if node is not None else default

    md = []
    md.append("# Top_Offenders_Customers - aplicação do ajuste")
    md.append("")
    md.append(f"Data/hora: {stamp}")
    md.append("")
    md.append("## Escopo")
    md.append("")
    md.append(f"- Aba: `{SHEET_NAME}`")
    md.append(f"- Tabela dinâmica: `{PIVOT_NAME}`")
    md.append(f"- Campo ajustado: `{CUSTOMER_FIELD}`")
    md.append("- Ajuste: limpar filtro manual do campo de cliente na visão geral.")
    md.append("- Não alterado: blocos de atraso/helpers, Vendors, Week_Overview e Region errors.")
    md.append("")
    md.append("## Resultado final")
    md.append("")
    md.append(f"- Total final da PivotTable1: **{val(['final_validation','official_total'])}**")
    md.append(f"- Total esperado Week_Overview DS: **{val(['final_validation','expected_week_overview_ds'])}**")
    md.append(f"- Bate com Week_Overview DS: **{val(['final_validation','matches_week_overview_ds'])}**")
    md.append(f"- Clientes visíveis/total: **{val(['final_validation','customer_visible'])}/{val(['final_validation','customer_items'])}**")
    md.append(f"- Clientes ocultos: **{val(['final_validation','customer_hidden'])}**")
    md.append("")
    md.append("## Arquivos")
    md.append("")
    md.append(f"- Cópia de teste: `{test_copy}`")
    md.append(f"- Backup oficial antes do ajuste: `{backup}`")
    md.append(f"- Relatório JSON: `{report_path}`")
    md_path.write_text("\n".join(md), encoding="utf-8")
    print(json.dumps(as_jsonable({"report": str(report_path), "summary": report["final_validation"]}), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()


