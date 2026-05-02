from __future__ import annotations

import re
import shutil
import zipfile
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
import xml.etree.ElementTree as ET

import openpyxl

try:
    from scripts.workbook_paths import resolve_workbook_path
except ModuleNotFoundError:
    from workbook_paths import resolve_workbook_path

WORKBOOK_PATH = resolve_workbook_path()
PROJECT_DIR = Path(__file__).resolve().parents[1]
BACKUP_DIR = PROJECT_DIR / "backups"

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_XR = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
NS_XR3 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"
NS = {"main": NS_MAIN}

ET.register_namespace("", NS_MAIN)
ET.register_namespace("r", NS_REL)
ET.register_namespace("mc", NS_MC)
ET.register_namespace("xr", NS_XR)
ET.register_namespace("xr3", NS_XR3)

OK = "Ok"
FALLBACK_SHEET_NAME = "PortRegion_Fallback"

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


def backup_workbook(path: Path) -> Path:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = BACKUP_DIR / f"{path.stem}_backup_{timestamp}{path.suffix}"
    shutil.copy2(path, backup_path)
    return backup_path


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


def col_num(ref: str) -> int:
    letters = re.match(r"[A-Z]+", ref).group(0)
    value = 0
    for char in letters:
        value = value * 26 + (ord(char) - 64)
    return value


def ensure_row(sheet_data: ET.Element, row_number: int) -> ET.Element:
    for row in sheet_data.findall("main:row", NS):
        if int(row.attrib["r"]) == row_number:
            return row
    row = ET.SubElement(sheet_data, f"{{{NS_MAIN}}}row", {"r": str(row_number)})
    return row


def get_cell(row: ET.Element, ref: str) -> ET.Element | None:
    for cell in row.findall("main:c", NS):
        if cell.attrib.get("r") == ref:
            return cell
    return None


def ensure_cell(row: ET.Element, ref: str, style: str | None = None) -> ET.Element:
    cell = get_cell(row, ref)
    if cell is None:
        attrs = {"r": ref}
        if style is not None:
            attrs["s"] = style
        cell = ET.Element(f"{{{NS_MAIN}}}c", attrs)
        row.append(cell)
    elif style is not None:
        cell.attrib["s"] = style
    return cell


def clear_cell(cell: ET.Element) -> None:
    for child in list(cell):
        cell.remove(child)
    cell.attrib.pop("t", None)


def set_inline_string(cell: ET.Element, value: str) -> None:
    clear_cell(cell)
    cell.attrib["t"] = "inlineStr"
    inline = ET.SubElement(cell, f"{{{NS_MAIN}}}is")
    text = ET.SubElement(inline, f"{{{NS_MAIN}}}t")
    text.text = value


def set_number(cell: ET.Element, value: int | float) -> None:
    clear_cell(cell)
    v = ET.SubElement(cell, f"{{{NS_MAIN}}}v")
    v.text = str(value)


def set_formula(cell: ET.Element, formula: str, *, result_type: str = "str") -> None:
    clear_cell(cell)
    if result_type:
        cell.attrib["t"] = result_type
    formula_node = ET.SubElement(cell, f"{{{NS_MAIN}}}f")
    formula_node.text = formula


def sort_row_cells(row: ET.Element) -> None:
    cells = list(row.findall("main:c", NS))
    for cell in cells:
        row.remove(cell)
    cells.sort(key=lambda item: col_num(item.attrib["r"]))
    for cell in cells:
        row.append(cell)


def patch_workbook_xml(root: ET.Element) -> None:
    calc = root.find("main:calcPr", NS)
    if calc is None:
        calc = ET.SubElement(root, f"{{{NS_MAIN}}}calcPr")
    calc.attrib["calcMode"] = "auto"
    calc.attrib["fullCalcOnLoad"] = "1"
    calc.attrib["forceFullCalc"] = "1"
    calc.attrib["calcCompleted"] = "0"

    sheets = root.find("main:sheets", NS)
    for sheet in sheets:
        if sheet.attrib.get("name") == "Planilha1":
            sheet.attrib["name"] = FALLBACK_SHEET_NAME
            break


def patch_fallback_sheet(root: ET.Element, assignments: list[dict[str, str | int]]) -> None:
    dim = root.find("main:dimension", NS)
    dim.attrib["ref"] = f"A1:I{len(assignments) + 1}"

    sheet_data = root.find("main:sheetData", NS)
    for row in list(sheet_data):
        sheet_data.remove(row)

    header_style = "1"
    data_style = "0"

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

    header_row = ensure_row(sheet_data, 1)
    for index, header in enumerate(headers, start=1):
        ref = f"{chr(64 + index)}1"
        set_inline_string(ensure_cell(header_row, ref, style=header_style), header)
    sort_row_cells(header_row)

    for row_number, item in enumerate(assignments, start=2):
        row = ensure_row(sheet_data, row_number)
        values = [
            item["os"] or "",
            item["port"] or "",
            item["region"] or "",
            item["match_status"] or "",
            item["reason"] or "",
            item["row"],
            item["booking"] or "",
            item["client"] or "",
            item["provider"] or "",
        ]
        for index, value in enumerate(values, start=1):
            ref = f"{chr(64 + index)}{row_number}"
            cell = ensure_cell(row, ref, style=data_style)
            if isinstance(value, int):
                set_number(cell, value)
            else:
                set_inline_string(cell, str(value))
        sort_row_cells(row)


def patch_roe_sheet(root: ET.Element) -> None:
    dim = root.find("main:dimension", NS)
    dim.attrib["ref"] = "A1:CI9915"

    sheet_data = root.find("main:sheetData", NS)

    header_row = ensure_row(sheet_data, 1)
    row1_style = get_cell(header_row, "CF1").attrib.get("s", "7") if get_cell(header_row, "CF1") is not None else "7"
    set_inline_string(ensure_cell(header_row, "CG1", style=row1_style), "Porto Base")
    set_inline_string(ensure_cell(header_row, "CH1", style=row1_style), "Region Base")
    set_inline_string(ensure_cell(header_row, "CI1", style=row1_style), "Match_Status")
    sort_row_cells(header_row)

    for row_number in range(2, 9916):
        row = ensure_row(sheet_data, row_number)

        sample_style = None
        for sample_ref in [f"AY{row_number}", f"AZ{row_number}", f"CF{row_number}"]:
            sample_cell = get_cell(row, sample_ref)
            if sample_cell is not None and sample_cell.attrib.get("s") is not None:
                sample_style = sample_cell.attrib["s"]
                break

        cg = ensure_cell(row, f"CG{row_number}", style=sample_style)
        ch = ensure_cell(row, f"CH{row_number}", style=sample_style)
        ci = ensure_cell(row, f"CI{row_number}", style=sample_style)
        ay = ensure_cell(row, f"AY{row_number}", style=sample_style)
        az = ensure_cell(row, f"AZ{row_number}", style=sample_style)

        set_formula(cg, f'IF(AW{row_number}<>"",AW{row_number},AX{row_number})')
        set_formula(ch, f'IFERROR(_xlfn.XLOOKUP(CG{row_number},Aux!E:E,Aux!D:D,""),"")')
        set_formula(ci, f'IF(AND(CG{row_number}<>"",CH{row_number}<>""),"RAO",IFERROR(_xlfn.XLOOKUP(A{row_number},{FALLBACK_SHEET_NAME}!A:A,{FALLBACK_SHEET_NAME}!D:D,""),"Sem Match"))')
        set_formula(ay, f'IF(CG{row_number}<>"",CG{row_number},IFERROR(_xlfn.XLOOKUP(A{row_number},{FALLBACK_SHEET_NAME}!A:A,{FALLBACK_SHEET_NAME}!B:B,""),""))')
        set_formula(az, f'IF(CH{row_number}<>"",CH{row_number},IFERROR(_xlfn.XLOOKUP(A{row_number},{FALLBACK_SHEET_NAME}!A:A,{FALLBACK_SHEET_NAME}!C:C,""),""))')

        sort_row_cells(row)


def patch_table3(root: ET.Element) -> None:
    for column in root.findall(".//main:tableColumn", NS):
        name = column.attrib.get("name")
        formula = column.find("main:calculatedColumnFormula", NS)
        if formula is None:
            continue
        if name == "Porto":
            formula.text = f'IF(CG2<>"",CG2,IFERROR(_xlfn.XLOOKUP(A2,{FALLBACK_SHEET_NAME}!A:A,{FALLBACK_SHEET_NAME}!B:B,""),""))'
        elif name == "Region":
            formula.text = f'IF(CH2<>"",CH2,IFERROR(_xlfn.XLOOKUP(A2,{FALLBACK_SHEET_NAME}!A:A,{FALLBACK_SHEET_NAME}!C:C,""),""))'


def serialize_xml(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def main() -> None:
    assignments = infer_missing_port_region(WORKBOOK_PATH)
    if not assignments:
        raise RuntimeError("No missing rows were inferred.")

    backup_path = backup_workbook(WORKBOOK_PATH)
    temp_path = WORKBOOK_PATH.with_suffix(".port-region.tmp.xlsm")

    with zipfile.ZipFile(WORKBOOK_PATH, "r") as source, zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            if info.filename == "xl/calcChain.xml":
                continue

            data = source.read(info.filename)

            if info.filename == "xl/workbook.xml":
                root = ET.fromstring(data)
                patch_workbook_xml(root)
                data = serialize_xml(root)

            elif info.filename == "xl/worksheets/sheet37.xml":
                root = ET.fromstring(data)
                patch_fallback_sheet(root, assignments)
                data = serialize_xml(root)

            elif info.filename == "xl/worksheets/sheet9.xml":
                root = ET.fromstring(data)
                patch_roe_sheet(root)
                data = serialize_xml(root)

            elif info.filename == "xl/tables/table3.xml":
                root = ET.fromstring(data)
                patch_table3(root)
                data = serialize_xml(root)

            target.writestr(info, data)

    shutil.move(temp_path, WORKBOOK_PATH)
    print(f"Backup created at: {backup_path}")
    print(f"Applied final port/region layer for {len(assignments)} rows.")


if __name__ == "__main__":
    main()
