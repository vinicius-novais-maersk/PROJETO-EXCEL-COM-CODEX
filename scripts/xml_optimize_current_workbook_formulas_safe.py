from __future__ import annotations

import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from scripts.update_special_clients import WORKBOOK_PATH, backup_workbook


NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
ET.register_namespace("", NS)

ROE_WK_SHEET = "xl/worksheets/sheet9.xml"
ROE_WK_MONTHLY_SHEET = "xl/worksheets/sheet27.xml"

BI_FORMULA = (
    'IF([@[Tipo Serviço]]="Transporte Rodoviário",'
    'IF(COUNTIFS(ROE_wk[Provedor],[@Provedor],ROE_wk[Tipo Serviço],[@[Tipo Serviço]],'
    'ROE_wk[Porto],[@Porto],ROE_wk[Weeknum],[@Weeknum],ROE_wk[Volume],"Ok",'
    'ROE_wk[OTO Out],"N",ROE_wk[OTD ajustado],"<>Sem Preenchimento",ROE_wk[Especiais],"<>Especial")=0,"",'
    '1-COUNTIFS(ROE_wk[OTD ajustado],"Atrasado",ROE_wk[Provedor],[@Provedor],'
    'ROE_wk[Tipo Serviço],[@[Tipo Serviço]],ROE_wk[Porto],[@Porto],ROE_wk[Weeknum],[@Weeknum],'
    'ROE_wk[Volume],"Ok",ROE_wk[OTO Out],"N",ROE_wk[Especiais],"<>Especial")/'
    'COUNTIFS(ROE_wk[Provedor],[@Provedor],ROE_wk[Tipo Serviço],[@[Tipo Serviço]],'
    'ROE_wk[Porto],[@Porto],ROE_wk[Weeknum],[@Weeknum],ROE_wk[Volume],"Ok",'
    'ROE_wk[OTO Out],"N",ROE_wk[OTD ajustado],"<>Sem Preenchimento",ROE_wk[Especiais],"<>Especial")),"")'
)

BJ_FORMULA = (
    'IF([@[Tipo Serviço]]="Transporte Rodoviário",'
    'IF(COUNTIFS(ROE_wk[Cliente Proposta],[@[Cliente Proposta]],ROE_wk[Tipo Serviço],[@[Tipo Serviço]],'
    'ROE_wk[Porto],[@Porto],ROE_wk[Weeknum],[@Weeknum],ROE_wk[Volume],"Ok",'
    'ROE_wk[OTO Out],"N",ROE_wk[OTD ajustado],"<>Sem Preenchimento",ROE_wk[Especiais],"<>Especial")=0,"",'
    '1-COUNTIFS(ROE_wk[OTD ajustado],"Atrasado",ROE_wk[Cliente Proposta],[@[Cliente Proposta]],'
    'ROE_wk[Tipo Serviço],[@[Tipo Serviço]],ROE_wk[Porto],[@Porto],ROE_wk[Weeknum],[@Weeknum],'
    'ROE_wk[Volume],"Ok",ROE_wk[OTO Out],"N",ROE_wk[Especiais],"<>Especial")/'
    'COUNTIFS(ROE_wk[Cliente Proposta],[@[Cliente Proposta]],ROE_wk[Tipo Serviço],[@[Tipo Serviço]],'
    'ROE_wk[Porto],[@Porto],ROE_wk[Weeknum],[@Weeknum],ROE_wk[Volume],"Ok",'
    'ROE_wk[OTO Out],"N",ROE_wk[OTD ajustado],"<>Sem Preenchimento",ROE_wk[Especiais],"<>Especial")),"")'
)

BK_FORMULA = (
    'IF([@[Tipo Serviço]]="Transporte Rodoviário",'
    'IF(COUNTIFS(ROE_wk[Cliente Proposta],[@[Cliente Proposta]],ROE_wk[Tipo Serviço],[@[Tipo Serviço]],'
    'ROE_wk[Porto],[@Porto],ROE_wk[Volume],"Ok")=0,"",'
    '1-COUNTIFS(ROE_wk[SLA Ag],"N",ROE_wk[Cliente Proposta],[@[Cliente Proposta]],'
    'ROE_wk[Tipo Serviço],[@[Tipo Serviço]],ROE_wk[Porto],[@Porto],ROE_wk[Volume],"Ok")/'
    'COUNTIFS(ROE_wk[Cliente Proposta],[@[Cliente Proposta]],ROE_wk[Tipo Serviço],[@[Tipo Serviço]],'
    'ROE_wk[Porto],[@Porto],ROE_wk[Volume],"Ok")),"")'
)


def patch_roe_wk(xml_bytes: bytes) -> tuple[bytes, dict[str, int]]:
    root = ET.fromstring(xml_bytes)
    counts = {"BI": 0, "BJ": 0, "BK": 0}
    for cell in root.iter(f"{{{NS}}}c"):
        ref = cell.attrib.get("r", "")
        formula = cell.find(f"{{{NS}}}f")
        if formula is None:
            continue
        if ref.startswith("BI"):
            formula.text = BI_FORMULA
            counts["BI"] += 1
        elif ref.startswith("BJ"):
            formula.text = BJ_FORMULA
            counts["BJ"] += 1
        elif ref.startswith("BK"):
            formula.text = BK_FORMULA
            counts["BK"] += 1
    return ET.tostring(root, encoding="utf-8", xml_declaration=False), counts


def patch_roe_wk_monthly(xml_bytes: bytes) -> tuple[bytes, int]:
    root = ET.fromstring(xml_bytes)
    count = 0
    for cell in root.iter(f"{{{NS}}}c"):
        ref = cell.attrib.get("r", "")
        if not ref.startswith("BJ"):
            continue
        formula = cell.find(f"{{{NS}}}f")
        if formula is None or not formula.text:
            continue
        row_num = "".join(ch for ch in ref if ch.isdigit())
        formula.text = f'XLOOKUP(A{row_num},ROE_wk[Nº OS],ROE_wk[% OTO Provider],"")'
        count += 1
    return ET.tostring(root, encoding="utf-8", xml_declaration=False), count


def main() -> int:
    if not WORKBOOK_PATH.exists():
        print(f"Workbook not found: {WORKBOOK_PATH}")
        return 1

    backup_path = backup_workbook(WORKBOOK_PATH)
    print(f"Backup created at: {backup_path}")

    with zipfile.ZipFile(WORKBOOK_PATH, "r") as zin:
        roe_wk_xml, roe_counts = patch_roe_wk(zin.read(ROE_WK_SHEET))
        roe_wk_monthly_xml, monthly_count = patch_roe_wk_monthly(zin.read(ROE_WK_MONTHLY_SHEET))

        with tempfile.NamedTemporaryFile(delete=False, suffix=WORKBOOK_PATH.suffix) as tmp:
            tmp_path = Path(tmp.name)

        try:
            with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == ROE_WK_SHEET:
                        data = roe_wk_xml
                    elif item.filename == ROE_WK_MONTHLY_SHEET:
                        data = roe_wk_monthly_xml
                    zout.writestr(item, data)

            shutil.move(tmp_path, WORKBOOK_PATH)
        finally:
            if tmp_path.exists():
                tmp_path.unlink(missing_ok=True)

    print(f"ROE_wk optimized cells: {roe_counts}")
    print(f"ROE_wk_monthly optimized BJ cells: {monthly_count}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
