from __future__ import annotations

import re
import shutil
import zipfile
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
FALLBACK_SHEET = "PortRegion_Fallback"

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


def backup_workbook(path: Path) -> Path:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = BACKUP_DIR / f"{path.stem}_backup_{timestamp}{path.suffix}"
    shutil.copy2(path, backup_path)
    return backup_path


def load_static_assignments(path: Path) -> list[dict[str, object]]:
    wb = openpyxl.load_workbook(path, keep_vba=True, data_only=True, read_only=True)
    ws = wb[FALLBACK_SHEET]
    assignments = []
    for row in ws.iter_rows(min_row=2, max_col=9, values_only=True):
        if not row[0]:
            continue
        assignments.append(
            {
                "os": row[0],
                "port": row[1],
                "region": row[2],
                "match_status": row[3],
                "reason": row[4],
                "roe_row": row[5],
                "booking": row[6],
                "client": row[7],
                "provider": row[8],
            }
        )
    return assignments


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


def col_num(ref: str) -> int:
    letters = re.match(r"[A-Z]+", ref).group(0)
    value = 0
    for char in letters:
        value = value * 26 + (ord(char) - 64)
    return value


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
    text.text = str(value)


def set_number(cell: ET.Element, value: int | float) -> None:
    clear_cell(cell)
    v = ET.SubElement(cell, f"{{{NS_MAIN}}}v")
    v.text = str(value)


def set_formula(cell: ET.Element, formula: str, *, result_type: str | None = None) -> None:
    clear_cell(cell)
    if result_type is not None:
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


def patch_workbook(root: ET.Element) -> None:
    calc = root.find("main:calcPr", NS)
    if calc is None:
        calc = ET.SubElement(root, f"{{{NS_MAIN}}}calcPr")
    calc.attrib["calcMode"] = "auto"
    calc.attrib["fullCalcOnLoad"] = "1"
    calc.attrib["forceFullCalc"] = "1"
    calc.attrib["calcCompleted"] = "0"


def patch_sheet37(root: ET.Element, assignments: list[dict[str, object]]) -> None:
    dim = root.find("main:dimension", NS)
    dim.attrib["ref"] = "A1:P250"

    cols = root.find("main:cols", NS)
    if cols is not None:
        root.remove(cols)
    cols = ET.Element(f"{{{NS_MAIN}}}cols")
    col_specs = [
        (1, 1, 18, False),
        (2, 2, 16, False),
        (3, 3, 14, False),
        (4, 4, 24, False),
        (5, 5, 12, False),
        (6, 6, 16, False),
        (7, 7, 22, False),
        (8, 8, 28, False),
        (9, 9, 46, False),
        (10, 16, 14, True),
    ]
    for min_col, max_col, width, hidden in col_specs:
        attrs = {
            "min": str(min_col),
            "max": str(max_col),
            "width": str(width),
            "customWidth": "1",
        }
        if hidden:
            attrs["hidden"] = "1"
        cols.append(ET.Element(f"{{{NS_MAIN}}}col", attrs))
    root.insert(3, cols)

    sheet_data = root.find("main:sheetData", NS)
    for row in list(sheet_data):
        sheet_data.remove(row)

    title_style = "7"
    label_style = "7"
    data_style = "0"

    visible_rows = {
        1: ["Fallback de Porto/Região - Casos Ativos"],
        2: ["Esta aba mostra apenas OS que ainda estão sem classificação oficial no RAO_wk."],
        3: ["Se a OS aparecer corretamente no RAO_wk, ela sai desta lista automaticamente."],
        4: ["A base técnica usada para inferência continua armazenada nesta aba, em colunas ocultas."],
        6: ["Casos ativos", None, "Sem Match", None, "Inferidos"],
        7: [
            '=COUNTIFS(ROE_wk!$A$2:$A$9915,"<>",ROE_wk!$CI$2:$CI$9915,"<>RAO")',
            None,
            '=COUNTIFS(ROE_wk!$A$2:$A$9915,"<>",ROE_wk!$CI$2:$CI$9915,"Sem Match")',
            None,
            '=COUNTIFS(ROE_wk!$A$2:$A$9915,"<>",ROE_wk!$CI$2:$CI$9915,"<>RAO",ROE_wk!$CI$2:$CI$9915,"<>Sem Match")',
        ],
        9: [
            "Nº OS",
            "Porto Final",
            "Region Final",
            "Match_Status",
            "Linha ROE_wk",
            "Booking",
            "Cliente",
            "Provedor",
            "Observação",
        ],
    }

    for row_number, values in visible_rows.items():
        row = ensure_row(sheet_data, row_number)
        for index, value in enumerate(values, start=1):
            if value is None:
                continue
            ref = f"{chr(64 + index)}{row_number}"
            style = title_style if row_number in {1, 2, 3, 4, 9, 6} else data_style
            cell = ensure_cell(row, ref, style=style)
            if isinstance(value, str) and value.startswith("="):
                set_formula(cell, value[1:])
            else:
                set_inline_string(cell, value)
        sort_row_cells(row)

    helper_formula = 'IFERROR(AGGREGATE(15,6,ROW(ROE_wk!$A$2:$A$9915)/((ROE_wk!$A$2:$A$9915<>"")*(ROE_wk!$CI$2:$CI$9915<>"RAO")),ROWS($J$10:J10)),"")'
    data_formulas = {
        "A": 'IF($J{r}="","",INDEX(ROE_wk!$A:$A,$J{r}))',
        "B": 'IF($J{r}="","",INDEX(ROE_wk!$AY:$AY,$J{r}))',
        "C": 'IF($J{r}="","",INDEX(ROE_wk!$AZ:$AZ,$J{r}))',
        "D": 'IF($J{r}="","",INDEX(ROE_wk!$CI:$CI,$J{r}))',
        "E": 'IF($J{r}="","",$J{r})',
        "F": 'IF($J{r}="","",INDEX(ROE_wk!$H:$H,$J{r}))',
        "G": 'IF($J{r}="","",INDEX(ROE_wk!$M:$M,$J{r}))',
        "H": 'IF($J{r}="","",INDEX(ROE_wk!$D:$D,$J{r}))',
        "I": 'IF($J{r}="","",IF($D{r}="Sem Match","Sem RAO e sem inferência disponível","Sem RAO; usando inferência até a base oficial chegar"))',
    }

    for row_number in range(10, 251):
        row = ensure_row(sheet_data, row_number)
        j_cell = ensure_cell(row, f"J{row_number}", style=data_style)
        set_formula(j_cell, helper_formula.replace("{r}", str(row_number)))
        for column, template in data_formulas.items():
            cell = ensure_cell(row, f"{column}{row_number}", style=data_style)
            set_formula(cell, template.format(r=row_number))
        sort_row_cells(row)

    hidden_headers = ["Nº OS", "Porto Inferido", "Region Inferida", "Match_Status", "Motivo", "Linha ROE_wk", "Booking", "Cliente", "Provedor"]
    row = ensure_row(sheet_data, 1)
    for index, header in enumerate(hidden_headers, start=12):
        ref = f"{chr(64 + index)}1"
        set_inline_string(ensure_cell(row, ref, style=label_style), header)
    sort_row_cells(row)

    for row_number, item in enumerate(assignments, start=2):
        row = ensure_row(sheet_data, row_number)
        values = [
            item["os"] or "",
            item["port"] or "",
            item["region"] or "",
            item["match_status"] or "",
            item["reason"] or "",
            item["roe_row"] or "",
            item["booking"] or "",
            item["client"] or "",
            item["provider"] or "",
        ]
        for index, value in enumerate(values, start=12):
            ref = f"{chr(64 + index)}{row_number}"
            cell = ensure_cell(row, ref, style=data_style)
            if isinstance(value, (int, float)):
                set_number(cell, value)
            else:
                set_inline_string(cell, value)
        sort_row_cells(row)

    auto_filter = root.find("main:autoFilter", NS)
    if auto_filter is None:
        auto_filter = ET.SubElement(root, f"{{{NS_MAIN}}}autoFilter")
    auto_filter.attrib["ref"] = "A9:I250"


def patch_sheet9(root: ET.Element) -> None:
    sheet_data = root.find("main:sheetData", NS)
    for row in sheet_data.findall("main:row", NS):
        row_number = int(row.attrib["r"])
        if row_number == 1:
            continue
        for cell in row.findall("main:c", NS):
            ref = cell.attrib.get("r", "")
            formula = cell.find("main:f", NS)
            if formula is None:
                continue
            if ref.startswith("AY"):
                formula.text = f'IF(CG{row_number}<>"",CG{row_number},IFERROR(_xlfn.XLOOKUP(A{row_number},{FALLBACK_SHEET}!L:L,{FALLBACK_SHEET}!M:M,""),""))'
                for child in list(cell):
                    if child.tag == f"{{{NS_MAIN}}}v":
                        cell.remove(child)
            elif ref.startswith("AZ"):
                formula.text = f'IF(CH{row_number}<>"",CH{row_number},IFERROR(_xlfn.XLOOKUP(A{row_number},{FALLBACK_SHEET}!L:L,{FALLBACK_SHEET}!N:N,""),""))'
                for child in list(cell):
                    if child.tag == f"{{{NS_MAIN}}}v":
                        cell.remove(child)
            elif ref.startswith("CI"):
                formula.text = f'IF(AND(CG{row_number}<>"",CH{row_number}<>""),"RAO",IFERROR(_xlfn.XLOOKUP(A{row_number},{FALLBACK_SHEET}!L:L,{FALLBACK_SHEET}!O:O,""),"Sem Match"))'
                for child in list(cell):
                    if child.tag == f"{{{NS_MAIN}}}v":
                        cell.remove(child)


def patch_table3(root: ET.Element) -> None:
    for col in root.findall(".//main:tableColumn", NS):
        name = col.attrib.get("name")
        formula = col.find("main:calculatedColumnFormula", NS)
        if formula is None:
            continue
        if name == "Porto":
            formula.text = f'IF(CG2<>"",CG2,IFERROR(_xlfn.XLOOKUP(A2,{FALLBACK_SHEET}!L:L,{FALLBACK_SHEET}!M:M,""),""))'
        elif name == "Region":
            formula.text = f'IF(CH2<>"",CH2,IFERROR(_xlfn.XLOOKUP(A2,{FALLBACK_SHEET}!L:L,{FALLBACK_SHEET}!N:N,""),""))'


def serialize_xml(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def main() -> None:
    assignments = load_static_assignments(WORKBOOK_PATH)
    if not assignments:
        raise RuntimeError("PortRegion_Fallback is empty; nothing to upgrade.")

    backup_path = backup_workbook(WORKBOOK_PATH)
    temp_path = WORKBOOK_PATH.with_suffix(".dynamic-fallback.tmp.xlsm")

    with zipfile.ZipFile(WORKBOOK_PATH, "r") as source, zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            if info.filename == "xl/calcChain.xml":
                continue
            data = source.read(info.filename)

            if info.filename == "xl/workbook.xml":
                root = ET.fromstring(data)
                patch_workbook(root)
                data = serialize_xml(root)
            elif info.filename == "xl/worksheets/sheet37.xml":
                root = ET.fromstring(data)
                patch_sheet37(root, assignments)
                data = serialize_xml(root)
            elif info.filename == "xl/worksheets/sheet9.xml":
                root = ET.fromstring(data)
                patch_sheet9(root)
                data = serialize_xml(root)
            elif info.filename == "xl/tables/table3.xml":
                root = ET.fromstring(data)
                patch_table3(root)
                data = serialize_xml(root)

            target.writestr(info, data)

    shutil.move(temp_path, WORKBOOK_PATH)
    print(f"Backup created at: {backup_path}")
    print("PortRegion_Fallback upgraded to dynamic view.")


if __name__ == "__main__":
    main()
