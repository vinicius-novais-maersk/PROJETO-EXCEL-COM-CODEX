from __future__ import annotations

import re
import shutil
import zipfile
from datetime import datetime
from pathlib import Path
import xml.etree.ElementTree as ET

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

ET.register_namespace("", NS_MAIN)
ET.register_namespace("r", NS_REL)
ET.register_namespace("mc", NS_MC)
ET.register_namespace("xr", NS_XR)
ET.register_namespace("xr3", NS_XR3)

NS = {"main": NS_MAIN}

MONTH_REGION_GROUPS = {
    4: [2, 3],
    9: [6, 7, 8],
    14: [11, 12, 13],
    20: [16, 17, 18, 19],
}

ALLIANCE = "Aliança"
OK = "Ok"
ATRASADO = "Atrasado"
SEM_PREENCHIMENTO = "Sem Preenchimento"
ESPECIAL = "Especial"
REAGENDA = "Reagenda"

TARGET_MONTH_SHEETS = {"xl/worksheets/sheet17.xml", "xl/worksheets/sheet18.xml"}
ERROR_ONLY_SHEETS = {"xl/worksheets/sheet13.xml", "xl/worksheets/sheet15.xml"}


def q(value: str) -> str:
    return f'"{value}"'


def countifs(*parts: str) -> str:
    return "COUNTIFS(" + ",".join(parts) + ")"


def sumifs(sum_range: str, *parts: str) -> str:
    return "SUMIFS(" + ",".join([sum_range, *parts]) + ")"


def criterion(range_ref: str, value: str) -> str:
    return f"{range_ref},{value}"


def sheet_range(sheet_name: str, address: str) -> str:
    return f"{sheet_name}!{address}"


def sum_exprs(expressions: list[str]) -> str:
    if not expressions:
        return "0"
    if len(expressions) == 1:
        return expressions[0]
    return "(" + ")+(".join(expressions) + ")"


def build_monthly_denominator(port_ref: str, *, alliance_filter: str | None = None) -> str:
    parts = [
        criterion(sheet_range("ROE_wk_monthly", "$AV:$AV"), q(OK)),
        criterion(sheet_range("ROE_wk_monthly", "$AY:$AY"), port_ref),
        criterion(sheet_range("ROE_wk_monthly", "$BG:$BG"), "$Q$1"),
        criterion(sheet_range("ROE_wk_monthly", "$AZ:$AZ"), q(f"<>{SEM_PREENCHIMENTO}")),
        criterion(sheet_range("ROE_wk_monthly", "$BI:$BI"), q(f"<>{ESPECIAL}")),
    ]
    if alliance_filter is not None:
        parts.append(criterion(sheet_range("ROE_wk_monthly", "$AI:$AI"), q(alliance_filter)))
    return countifs(*parts)


def build_monthly_late(port_ref: str, *, alliance_filter: str | None = None) -> str:
    parts = [
        criterion(sheet_range("ROE_wk_monthly", "$AV:$AV"), q(OK)),
        criterion(sheet_range("ROE_wk_monthly", "$AY:$AY"), port_ref),
        criterion(sheet_range("ROE_wk_monthly", "$BG:$BG"), "$Q$1"),
        criterion(sheet_range("ROE_wk_monthly", "$BI:$BI"), q(f"<>{ESPECIAL}")),
        criterion(sheet_range("ROE_wk_monthly", "$AZ:$AZ"), q(ATRASADO)),
    ]
    if alliance_filter is not None:
        parts.append(criterion(sheet_range("ROE_wk_monthly", "$AI:$AI"), q(alliance_filter)))
    return countifs(*parts)


def build_monthly_schedule_denominator(port_ref: str) -> str:
    return countifs(
        criterion(sheet_range("ROE_wk_monthly", "$AV:$AV"), q(OK)),
        criterion(sheet_range("ROE_wk_monthly", "$AY:$AY"), port_ref),
        criterion(sheet_range("ROE_wk_monthly", "$BG:$BG"), "$Q$1"),
        criterion(sheet_range("ROE_wk_monthly", "$AI:$AI"), q(ALLIANCE)),
    )


def build_monthly_schedule_late(port_ref: str) -> str:
    return countifs(
        criterion(sheet_range("ROE_wk_monthly", "$AV:$AV"), q(OK)),
        criterion(sheet_range("ROE_wk_monthly", "$AY:$AY"), port_ref),
        criterion(sheet_range("ROE_wk_monthly", "$BG:$BG"), "$Q$1"),
        criterion(sheet_range("ROE_wk_monthly", "$AI:$AI"), q(ALLIANCE)),
        criterion(sheet_range("ROE_wk_monthly", "$AT:$AT"), q("N")),
    )


def build_monthly_reagenda_formula(port_ref: str) -> str:
    start_date = "DATE(YEAR($R$1),$Q$1,1)"
    end_date = "EOMONTH(DATE(YEAR($R$1),$Q$1,1),0)+1"
    return (
        "IF("
        + f"{port_ref}=\"\",\"\","
        + sumifs(
            sheet_range("Reagendas", "$Q:$Q"),
            criterion(sheet_range("Reagendas", "$F:$F"), port_ref),
            criterion(sheet_range("Reagendas", "$K:$K"), q(REAGENDA)),
            criterion(sheet_range("Reagendas", "$J:$J"), q(">=") + "&" + start_date),
            criterion(sheet_range("Reagendas", "$J:$J"), q("<") + "&" + end_date),
        )
        + ")"
    )


def build_monthly_hypercare(port_ref: str) -> str:
    return (
        "IF("
        + f"{port_ref}=\"\",\"\","
        + countifs(
            criterion(sheet_range("ROE_wk_monthly", "$AV:$AV"), q(OK)),
            criterion(sheet_range("ROE_wk_monthly", "$AY:$AY"), port_ref),
            criterion(sheet_range("ROE_wk_monthly", "$BG:$BG"), "$Q$1"),
            criterion(sheet_range("ROE_wk_monthly", "$AI:$AI"), q(ALLIANCE)),
            criterion(sheet_range("ROE_wk_monthly", "$N:$N"), q("*Reagenda*")),
        )
        + ")"
    )


def build_month_ratio_from_ports(child_rows: list[int], column: str) -> str:
    port_refs = [f"$Q${row}" for row in child_rows]
    if column == "N":
        denominator_exprs = [build_monthly_schedule_denominator(port_ref) for port_ref in port_refs]
        late_exprs = [build_monthly_schedule_late(port_ref) for port_ref in port_refs]
    else:
        alliance_filter = None
        if column == "K":
            alliance_filter = ALLIANCE
        elif column == "L":
            alliance_filter = f"<>{ALLIANCE}"
        denominator_exprs = [build_monthly_denominator(port_ref, alliance_filter=alliance_filter) for port_ref in port_refs]
        late_exprs = [build_monthly_late(port_ref, alliance_filter=alliance_filter) for port_ref in port_refs]
    denominator = sum_exprs(denominator_exprs)
    late = sum_exprs(late_exprs)
    return f'IF(({denominator})=0,"",1-IFERROR(({late})/({denominator}),0))'


def ratio_formula(port_ref: str, denominator: str, late: str) -> str:
    return f'IF({port_ref}="","",IF({denominator}=0,"",1-IFERROR({late}/{denominator},0)))'


def get_cell(root: ET.Element, ref: str) -> ET.Element | None:
    for cell in root.findall(".//main:c", NS):
        if cell.attrib.get("r") == ref:
            return cell
    return None


def strip_cached_value(cell: ET.Element, *, keep_type: str | None = None) -> None:
    for tag in ["v", "is"]:
        node = cell.find(f"main:{tag}", NS)
        if node is not None:
            cell.remove(node)
    if keep_type is None:
        cell.attrib.pop("t", None)
    else:
        cell.attrib["t"] = keep_type


def set_inline_string(cell: ET.Element, value: str) -> None:
    strip_cached_value(cell, keep_type="inlineStr")
    inline = cell.find("main:is", NS)
    if inline is None:
        inline = ET.SubElement(cell, f"{{{NS_MAIN}}}is")
    for child in list(inline):
        inline.remove(child)
    text = ET.SubElement(inline, f"{{{NS_MAIN}}}t")
    text.text = value


def set_formula(cell: ET.Element, formula: str, *, keep_type: str | None = None) -> None:
    formula_node = cell.find("main:f", NS)
    if formula_node is None:
        formula_node = ET.SubElement(cell, f"{{{NS_MAIN}}}f")
    formula_node.text = formula
    strip_cached_value(cell, keep_type=keep_type)


def delete_cell(root: ET.Element, ref: str) -> None:
    for row in root.findall(".//main:row", NS):
        for cell in list(row.findall("main:c", NS)):
            if cell.attrib.get("r") == ref:
                row.remove(cell)
                return


def row_number(ref: str) -> int:
    return int(re.search(r"(\d+)$", ref).group(1))


def patch_month_sheet(root: ET.Element) -> None:
    for row, child_rows in MONTH_REGION_GROUPS.items():
        set_formula(get_cell(root, f"K{row}"), build_month_ratio_from_ports(child_rows, "K"))
        set_formula(get_cell(root, f"L{row}"), build_month_ratio_from_ports(child_rows, "L"))
        set_formula(get_cell(root, f"M{row}"), build_month_ratio_from_ports(child_rows, "M"))
        set_formula(get_cell(root, f"N{row}"), build_month_ratio_from_ports(child_rows, "N"))
        set_formula(get_cell(root, f"O{row}"), f"SUM(O{child_rows[0]}:O{child_rows[-1]})")
        set_formula(get_cell(root, f"P{row}"), f"SUM(P{child_rows[0]}:P{child_rows[-1]})")

    for cell_ref in ["K22", "L22", "M22", "N22", "O22", "P22"]:
        delete_cell(root, cell_ref)


def patch_roe_weekly(root: ET.Element) -> None:
    for cell in root.findall(".//main:c", NS):
        ref = cell.attrib.get("r", "")
        formula = cell.find("main:f", NS)
        if ref.startswith("BH"):
            if ref == "BH1":
                set_inline_string(cell, "W8ngMinute")
                continue
            if formula is not None and "ref" in formula.attrib:
                row = row_number(ref)
                formula.text = f'IFERROR(BG{row}*24*60,"")'
            strip_cached_value(cell)
        elif ref.startswith("BQ") and formula is not None and formula.text and "VLOOKUP(" in formula.text and "IFERROR(" not in formula.text:
            formula.text = (
                'IF(ROE_wk[[#This Row],[OTD ajustado]]="Atrasado",'
                'IFERROR(VLOOKUP(ROE_wk[[#This Row],[Justificativa]],Table3[],2,FALSE),""),"")'
            )
            strip_cached_value(cell, keep_type="str")


def patch_roe_weekly_cancel(root: ET.Element) -> None:
    for cell in root.findall(".//main:c", NS):
        ref = cell.attrib.get("r", "")
        if not ref.startswith("BF"):
            continue
        if ref == "BF1":
            set_inline_string(cell, "W8ngMinute")
            continue
        formula = cell.find("main:f", NS)
        if formula is not None and "ref" in formula.attrib:
            row = row_number(ref)
            formula.text = f'IFERROR(BE{row}*24*60,"")'
        strip_cached_value(cell)


def patch_roe_monthly(root: ET.Element) -> None:
    for cell in root.findall(".//main:c", NS):
        ref = cell.attrib.get("r", "")
        formula = cell.find("main:f", NS)
        if formula is not None and formula.text:
            formula.text = formula.text.replace("Aux!$A$1:$A$6", "Aux!$A$1:$A$7").replace("Aux!$B$1:$B$6", "Aux!$B$1:$B$7")

        if ref.startswith("BF"):
            if ref == "BF1":
                set_inline_string(cell, "W8ngMinute")
                continue
            if formula is not None and "ref" in formula.attrib:
                row = row_number(ref)
                formula.text = f'IFERROR(BE{row}*24*60,"")'
            strip_cached_value(cell)
        elif ref.startswith("BH") and formula is not None and formula.text:
            if not formula.text.startswith("IFERROR("):
                formula.text = f'IFERROR({formula.text},"N")'
            strip_cached_value(cell, keep_type="str")


def patch_reagendas_sheet(root: ET.Element) -> None:
    for cell in root.findall(".//main:c", NS):
        formula = cell.find("main:f", NS)
        if formula is None or not formula.text:
            continue
        formula.text = formula.text.replace("Aux!$A$1:$A$6", "Aux!$A$1:$A$7").replace("Aux!$B$1:$B$6", "Aux!$B$1:$B$7")
        ref = cell.attrib.get("r", "")
        if ref == "P514" and "TEXTJOIN" in formula.text and not formula.text.startswith("IFERROR("):
            formula.text = f'IFERROR({formula.text},"")'
            strip_cached_value(cell, keep_type="str")
        elif ref == "Q514":
            formula.text = 'IF(P514="",0,LEN(P514)-LEN(SUBSTITUTE(P514,CHAR(10),""))+1)'
            strip_cached_value(cell)


def patch_support_errors(root: ET.Element) -> None:
    for row in root.findall(".//main:row", NS):
        for cell in list(row.findall("main:c", NS)):
            if cell.attrib.get("t") == "e" and cell.find("main:f", NS) is None:
                row.remove(cell)


def patch_table21(root: ET.Element) -> None:
    for calc in root.findall(".//main:calculatedColumnFormula", NS):
        if calc.text is None:
            continue
        calc.text = calc.text.replace("Aux!$A$1:$A$6", "Aux!$A$1:$A$7").replace("Aux!$B$1:$B$6", "Aux!$B$1:$B$7")
        if "TEXTJOIN(" in calc.text and not calc.text.startswith("IFERROR("):
            calc.text = f'IFERROR({calc.text},"")'
        elif calc.text == 'LEN(P2) - LEN(SUBSTITUTE(P2,CHAR(10),"")) + 1':
            calc.text = 'IF(P2="",0,LEN(P2)-LEN(SUBSTITUTE(P2,CHAR(10),""))+1)'


def patch_table3(root: ET.Element) -> None:
    for col in root.findall(".//main:tableColumn", NS):
        name = col.attrib.get("name")
        calc = col.find("main:calculatedColumnFormula", NS)
        if calc is None or calc.text is None:
            continue
        if name == "W8ngMinute":
            calc.text = 'IFERROR(BG2*24*60,"")'
        elif name == "Tipo Atraso":
            calc.text = (
                'IF(ROE_wk[[#This Row],[OTD ajustado]]="Atrasado",'
                'IFERROR(VLOOKUP(ROE_wk[[#This Row],[Justificativa]],Table3[],2,FALSE),""),"")'
            )


def patch_table20(root: ET.Element) -> None:
    for col in root.findall(".//main:tableColumn", NS):
        name = col.attrib.get("name")
        calc = col.find("main:calculatedColumnFormula", NS)
        if calc is None or calc.text is None:
            continue
        if name == "W8ngMinute":
            calc.text = 'IFERROR(BE2*24*60,"")'


def patch_workbook_xml(root: ET.Element) -> None:
    calc = root.find("main:calcPr", NS)
    if calc is None:
        calc = ET.SubElement(root, f"{{{NS_MAIN}}}calcPr")
    calc.attrib["calcMode"] = "auto"
    calc.attrib["fullCalcOnLoad"] = "1"
    calc.attrib["forceFullCalc"] = "1"
    calc.attrib["calcCompleted"] = "0"


def backup_workbook(path: Path) -> Path:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = BACKUP_DIR / f"{path.stem}_backup_{timestamp}{path.suffix}"
    shutil.copy2(path, backup_path)
    return backup_path


def serialize_xml(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def main() -> None:
    if not WORKBOOK_PATH.exists():
        raise FileNotFoundError(WORKBOOK_PATH)

    backup_path = backup_workbook(WORKBOOK_PATH)
    temp_path = WORKBOOK_PATH.with_suffix(".tmp.xlsm")

    with zipfile.ZipFile(WORKBOOK_PATH, "r") as source, zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            data = source.read(info.filename)

            if info.filename == "xl/calcChain.xml":
                continue

            if info.filename == "xl/workbook.xml":
                root = ET.fromstring(data)
                patch_workbook_xml(root)
                data = serialize_xml(root)

            elif info.filename in TARGET_MONTH_SHEETS:
                root = ET.fromstring(data)
                patch_month_sheet(root)
                data = serialize_xml(root)

            elif info.filename == "xl/worksheets/sheet9.xml":
                root = ET.fromstring(data)
                patch_roe_weekly(root)
                data = serialize_xml(root)

            elif info.filename == "xl/worksheets/sheet25.xml":
                root = ET.fromstring(data)
                patch_roe_weekly_cancel(root)
                data = serialize_xml(root)

            elif info.filename == "xl/worksheets/sheet27.xml":
                root = ET.fromstring(data)
                patch_roe_monthly(root)
                data = serialize_xml(root)

            elif info.filename == "xl/worksheets/sheet26.xml":
                root = ET.fromstring(data)
                patch_reagendas_sheet(root)
                data = serialize_xml(root)

            elif info.filename in ERROR_ONLY_SHEETS:
                root = ET.fromstring(data)
                patch_support_errors(root)
                data = serialize_xml(root)

            elif info.filename == "xl/tables/table21.xml":
                root = ET.fromstring(data)
                patch_table21(root)
                data = serialize_xml(root)

            elif info.filename == "xl/tables/table3.xml":
                root = ET.fromstring(data)
                patch_table3(root)
                data = serialize_xml(root)

            elif info.filename == "xl/tables/table20.xml":
                root = ET.fromstring(data)
                patch_table20(root)
                data = serialize_xml(root)

            if info.filename.endswith(".xml"):
                data = data.replace(b"Aux!$A$1:$A$6", b"Aux!$A$1:$A$7").replace(b"Aux!$B$1:$B$6", b"Aux!$B$1:$B$7")

            target.writestr(info, data)

    shutil.move(temp_path, WORKBOOK_PATH)
    print(f"Backup created at: {backup_path}")
    print(f"Workbook patched: {WORKBOOK_PATH}")


if __name__ == "__main__":
    main()
