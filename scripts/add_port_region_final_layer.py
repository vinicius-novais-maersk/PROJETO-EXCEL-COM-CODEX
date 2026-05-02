from __future__ import annotations

from collections import Counter, defaultdict
from pathlib import Path

import openpyxl
import pythoncom
import win32com.client

try:
    from scripts.workbook_paths import resolve_workbook_path
except ModuleNotFoundError:
    from workbook_paths import resolve_workbook_path

ROOT_DIR = Path(__file__).resolve().parents[1]
WORKBOOK_PATH = resolve_workbook_path()
FALLBACK_SHEET = "PortRegion_Fallback"

OK = "Ok"
PORT_TO_REGION = {
    "Manaus": "North",
    "Vila do Conde": "North",
    "Pecem": "Northeast",
    "Suape": "Northeast",
    "Salvador": "Northeast",
    "Santos": "Southeast",
    "Rio": "Southeast",
    "Vitoria": "Southeast",
    "Itapoa": "South",
    "Imbituba": "South",
    "Rio Grande": "South",
    "Itajai": "South",
    "Paranagua": "South",
    "Navegantes": "South",
}

MATCH_STATUS_LABELS = {
    "booking": "Inferido Booking",
    "client_provider": "Inferido Cliente+Provedor",
    "provider": "Inferido Provedor",
}


def infer_missing_port_region(workbook_path: Path) -> list[dict[str, str | int]]:
    wb = openpyxl.load_workbook(workbook_path, keep_vba=True, data_only=True, read_only=True)
    week = wb["Week_Overview"]["AG1"].value
    ws = wb["ROE_wk"]

    populated_rows = []
    missing_rows = []
    for row_index, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if row[47] != OK or row[43] != week:
            continue
        entry = {
            "row": row_index,
            "os": row[0],
            "booking": row[7],
            "client": row[12],
            "provider": row[3],
            "port": row[50],
            "region": row[51],
        }
        if row[50] in (None, "") or row[51] in (None, ""):
            missing_rows.append(entry)
        else:
            populated_rows.append(entry)

    by_booking: dict[object, list[str]] = defaultdict(list)
    by_client_provider: dict[tuple[object, object], list[str]] = defaultdict(list)
    by_provider: dict[object, list[str]] = defaultdict(list)

    for entry in populated_rows:
        by_booking[entry["booking"]].append(str(entry["port"]))
        by_client_provider[(entry["client"], entry["provider"])].append(str(entry["port"]))
        by_provider[entry["provider"]].append(str(entry["port"]))

    assignments = []
    for entry in missing_rows:
        candidates = Counter(by_booking.get(entry["booking"], []))
        reason = "booking"
        if len(candidates) != 1:
            candidates = Counter(by_client_provider.get((entry["client"], entry["provider"]), []))
            reason = "client_provider"
        if len(candidates) != 1:
            candidates = Counter(by_provider.get(entry["provider"], []))
            reason = "provider"
        if len(candidates) != 1:
            continue

        port = next(iter(candidates))
        region = PORT_TO_REGION.get(port)
        if region is None:
            continue

        assignments.append(
            {
                "row": entry["row"],
                "os": entry["os"],
                "booking": entry["booking"],
                "client": entry["client"],
                "provider": entry["provider"],
                "port": port,
                "region": region,
                "reason": reason,
                "match_status": MATCH_STATUS_LABELS[reason],
            }
        )

    assignments.sort(key=lambda item: int(item["row"]))
    return assignments


def ensure_sheet(workbook, sheet_name: str, after_sheet_name: str):
    for worksheet in workbook.Worksheets:
        if worksheet.Name == sheet_name:
            worksheet.Cells.Clear()
            return worksheet
    after_sheet = workbook.Worksheets(after_sheet_name)
    worksheet = workbook.Worksheets.Add(After=after_sheet)
    worksheet.Name = sheet_name
    return worksheet


def write_fallback_sheet(worksheet, assignments: list[dict[str, str | int]]) -> None:
    headers = [
        "Nº OS",
        "Porto Inferido",
        "Region Inferida",
        "Match_Status",
        "Motivo",
        "Linha ROE_wk",
        "Booking",
        "Cliente",
        "Provedor",
    ]
    for column_index, header in enumerate(headers, start=1):
        worksheet.Cells(1, column_index).Value = header

    for row_index, item in enumerate(assignments, start=2):
        worksheet.Cells(row_index, 1).Value = item["os"]
        worksheet.Cells(row_index, 2).Value = item["port"]
        worksheet.Cells(row_index, 3).Value = item["region"]
        worksheet.Cells(row_index, 4).Value = item["match_status"]
        worksheet.Cells(row_index, 5).Value = item["reason"]
        worksheet.Cells(row_index, 6).Value = item["row"]
        worksheet.Cells(row_index, 7).Value = item["booking"]
        worksheet.Cells(row_index, 8).Value = item["client"]
        worksheet.Cells(row_index, 9).Value = item["provider"]

    worksheet.Columns("A:I").AutoFit()


def main() -> None:
    assignments = infer_missing_port_region(WORKBOOK_PATH)
    if not assignments:
        raise RuntimeError("No missing port/region assignments were inferred.")

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    workbook = None

    try:
        workbook = excel.Workbooks.Open(str(WORKBOOK_PATH), UpdateLinks=0, ReadOnly=False)

        fallback_ws = ensure_sheet(workbook, FALLBACK_SHEET, "ROE_wk")
        write_fallback_sheet(fallback_ws, assignments)

        roe_ws = workbook.Worksheets("ROE_wk")
        last_row = roe_ws.Cells(roe_ws.Rows.Count, 1).End(-4162).Row

        roe_ws.Range("CG1").Value = "Porto Base"
        roe_ws.Range("CH1").Value = "Region Base"
        roe_ws.Range("CI1").Value = "Match_Status"

        roe_ws.Range(f"CG2:CG{last_row}").FormulaLocal = '=SE(AW2<>"";AW2;AX2)'
        roe_ws.Range(f"CH2:CH{last_row}").FormulaLocal = '=SEERRO(PROCX(CG2;Aux!E:E;Aux!D:D;"");"")'
        roe_ws.Range(f"CI2:CI{last_row}").FormulaLocal = f'=SE(E(CG2<>"";CH2<>"");"RAO";SEERRO(PROCX(A2;{FALLBACK_SHEET}!$A:$A;{FALLBACK_SHEET}!$D:$D;"");"Sem Match"))'

        list_object = roe_ws.ListObjects("ROE_wk")
        list_object.ListColumns("Porto").DataBodyRange.FormulaLocal = f'=SE(CG2<>"";CG2;SEERRO(PROCX(A2;{FALLBACK_SHEET}!$A:$A;{FALLBACK_SHEET}!$B:$B;"");""))'
        list_object.ListColumns("Region").DataBodyRange.FormulaLocal = f'=SE(CH2<>"";CH2;SEERRO(PROCX(A2;{FALLBACK_SHEET}!$A:$A;{FALLBACK_SHEET}!$C:$C;"");""))'

        roe_ws.Calculate()

        workbook.Save()
        print(f"Applied final port/region layer for {len(assignments)} rows.")
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
