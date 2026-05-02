from __future__ import annotations

import os
import shutil
import sys
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from scripts.update_special_clients import WORKBOOK_PATH, backup_workbook


NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
ET.register_namespace("", NS)

SHEET_XML = "xl/worksheets/sheet9.xml"
TABLE_XML = "xl/tables/table3.xml"
WORKBOOK_XML = "xl/workbook.xml"
WORKBOOK_RELS_XML = "xl/_rels/workbook.xml.rels"
CALC_CHAIN_XML = "xl/calcChain.xml"

DATA_START_ROW = 2
DATA_END_ROW = 9915

HELPER_COLUMN_NAME = "Embarcador+Cliente Proposta+Provedor+PortoExc"
MATCH_TYPE = "Embarcador+Cliente Proposta+Provedor+Porto"


def log(message: str) -> None:
    print(f"[procter-rio-irb-oto-xml] {message}", flush=True)


def qname(tag: str) -> str:
    return f"{{{NS}}}{tag}"


def rel_qname(tag: str) -> str:
    return f"{{{REL_NS}}}{tag}"


def split_cell_ref(ref: str) -> tuple[str, int]:
    letters = []
    digits = []
    for char in ref:
        if char.isalpha():
            letters.append(char)
        else:
            digits.append(char)
    return "".join(letters), int("".join(digits))


def column_index(col_ref: str) -> int:
    value = 0
    for char in col_ref:
        value = (value * 26) + ord(char.upper()) - 64
    return value


def get_row(sheet_data, row_number: int):
    for row in sheet_data.findall(qname("row")):
        if int(row.attrib["r"]) == row_number:
            return row
    row = ET.Element(qname("row"), {"r": str(row_number)})
    inserted = False
    for index, existing in enumerate(list(sheet_data)):
        if int(existing.attrib["r"]) > row_number:
            sheet_data.insert(index, row)
            inserted = True
            break
    if not inserted:
        sheet_data.append(row)
    return row


def get_or_create_cell(row, cell_ref: str, style: str | None = None):
    target_col, _ = split_cell_ref(cell_ref)
    target_idx = column_index(target_col)

    for cell in row.findall(qname("c")):
        col_ref, _ = split_cell_ref(cell.attrib["r"])
        if col_ref == target_col:
            return cell

    cell = ET.Element(qname("c"), {"r": cell_ref})
    if style is not None:
        cell.attrib["s"] = style

    inserted = False
    for index, existing in enumerate(list(row)):
        col_ref, _ = split_cell_ref(existing.attrib["r"])
        if column_index(col_ref) > target_idx:
            row.insert(index, cell)
            inserted = True
            break
    if not inserted:
        row.append(cell)
    return cell


def set_formula_cell(cell, formula: str, value_type: str | None) -> None:
    if value_type is None:
        cell.attrib.pop("t", None)
    else:
        cell.attrib["t"] = value_type

    for child in list(cell):
        if child.tag in {qname("f"), qname("v"), qname("is")}:
            cell.remove(child)

    formula_node = ET.Element(qname("f"))
    formula_node.text = formula
    cell.append(formula_node)


def get_formula(cell) -> str:
    formula_node = cell.find(qname("f"))
    if formula_node is None or not formula_node.text:
        raise RuntimeError(f"Missing formula in {cell.attrib['r']}.")
    return formula_node.text


def update_sheet_formulas(sheet_root) -> dict[str, str]:
    sheet_data = sheet_root.find(qname("sheetData"))
    if sheet_data is None:
        raise RuntimeError("sheetData not found in ROE_wk sheet XML.")

    template_cells = {}
    for ref in ("CA2", "CB2", "CC2"):
        row = get_row(sheet_data, 2)
        cell = get_or_create_cell(row, ref)
        template_cells[ref] = cell

    formulas = {
        "CA": get_formula(template_cells["CA2"]),
        "CB": get_formula(template_cells["CB2"]),
        "CC": get_formula(template_cells["CC2"]),
    }

    styles = {}
    for ref in ("CA3", "CB3", "CC2"):
        col, row_number = split_cell_ref(ref)
        row = get_row(sheet_data, row_number)
        cell = get_or_create_cell(row, ref)
        styles[col] = cell.attrib.get("s")

    for row_number in range(DATA_START_ROW, DATA_END_ROW + 1):
        row = get_row(sheet_data, row_number)
        for col, formula, value_type in (
            ("CA", formulas["CA"], "b"),
            ("CB", formulas["CB"], "b"),
            ("CC", formulas["CC"], "str"),
        ):
            cell_ref = f"{col}{row_number}"
            cell = get_or_create_cell(row, cell_ref, style=styles.get(col))
            if "s" not in cell.attrib and styles.get(col) is not None:
                cell.attrib["s"] = styles[col]
            set_formula_cell(cell, formula, value_type)

    return formulas


def update_table_formulas(table_root, formulas: dict[str, str]) -> None:
    table_columns = table_root.find(qname("tableColumns"))
    if table_columns is None:
        raise RuntimeError("tableColumns not found in ROE_wk table XML.")

    desired = {
        "Is_OTOException": formulas["CA"],
        "Is_MoveException": formulas["CB"],
        HELPER_COLUMN_NAME: formulas["CC"],
    }

    for column in table_columns.findall(qname("tableColumn")):
        name = column.attrib.get("name")
        if name not in desired:
            continue

        formula_node = column.find(qname("calculatedColumnFormula"))
        if formula_node is None:
            formula_node = ET.SubElement(column, qname("calculatedColumnFormula"))
        formula_node.text = desired[name]


def update_workbook_calc(workbook_root) -> None:
    calc_pr = workbook_root.find(qname("calcPr"))
    if calc_pr is None:
        calc_pr = ET.SubElement(workbook_root, qname("calcPr"))
    calc_pr.attrib["fullCalcOnLoad"] = "1"
    calc_pr.attrib["forceFullCalc"] = "1"
    calc_pr.attrib["calcCompleted"] = "0"


def remove_calc_chain(workbook_rels_root) -> None:
    for rel in list(workbook_rels_root):
        if rel.attrib.get("Type", "").endswith("/calcChain") or rel.attrib.get("Target") == "calcChain.xml":
            workbook_rels_root.remove(rel)


def rewrite_workbook(modified_parts: dict[str, bytes]) -> None:
    temp_fd, temp_name = tempfile.mkstemp(suffix=".xlsm", prefix="procter_rio_irb_")
    temp_path = Path(temp_name)
    os.close(temp_fd)
    try:
        with ZipFile(WORKBOOK_PATH, "r") as src, ZipFile(temp_path, "w") as dst:
            for info in src.infolist():
                if info.filename == CALC_CHAIN_XML:
                    continue

                payload = modified_parts.get(info.filename)
                if payload is None:
                    payload = src.read(info.filename)

                new_info = ZipInfo(info.filename, date_time=info.date_time)
                new_info.compress_type = info.compress_type or ZIP_DEFLATED
                new_info.comment = info.comment
                new_info.create_system = info.create_system
                new_info.create_version = info.create_version
                new_info.extract_version = info.extract_version
                new_info.flag_bits = info.flag_bits
                new_info.external_attr = info.external_attr
                new_info.internal_attr = info.internal_attr
                dst.writestr(new_info, payload)

        shutil.move(str(temp_path), str(WORKBOOK_PATH))
    finally:
        if temp_path.exists():
            temp_path.unlink(missing_ok=True)


def main() -> int:
    if not WORKBOOK_PATH.exists():
        log(f"Workbook not found: {WORKBOOK_PATH}")
        return 1

    backup_path = backup_workbook(WORKBOOK_PATH)
    log(f"Backup created at: {backup_path}")

    with ZipFile(WORKBOOK_PATH, "r") as workbook_zip:
        sheet_root = ET.fromstring(workbook_zip.read(SHEET_XML))
        table_root = ET.fromstring(workbook_zip.read(TABLE_XML))
        workbook_root = ET.fromstring(workbook_zip.read(WORKBOOK_XML))
        workbook_rels_root = ET.fromstring(workbook_zip.read(WORKBOOK_RELS_XML))

    formulas = update_sheet_formulas(sheet_root)
    if MATCH_TYPE not in formulas["CA"] or MATCH_TYPE not in formulas["CB"]:
        log("Template formulas do not contain the new match type. Aborting.")
        return 1

    update_table_formulas(table_root, formulas)
    update_workbook_calc(workbook_root)
    remove_calc_chain(workbook_rels_root)

    modified_parts = {
        SHEET_XML: ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True),
        TABLE_XML: ET.tostring(table_root, encoding="utf-8", xml_declaration=True),
        WORKBOOK_XML: ET.tostring(workbook_root, encoding="utf-8", xml_declaration=True),
        WORKBOOK_RELS_XML: ET.tostring(workbook_rels_root, encoding="utf-8", xml_declaration=True),
    }

    rewrite_workbook(modified_parts)
    log("Workbook XML updated successfully.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
