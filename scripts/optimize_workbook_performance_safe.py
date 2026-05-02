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


def main() -> int:
    if not WORKBOOK_PATH.exists():
        print(f"Workbook not found: {WORKBOOK_PATH}")
        return 1

    backup_path = backup_workbook(WORKBOOK_PATH)
    print(f"Backup created at: {backup_path}")

    with zipfile.ZipFile(WORKBOOK_PATH, "r") as zin:
        workbook_xml = ET.fromstring(zin.read("xl/workbook.xml"))
        calc_pr = workbook_xml.find(f"{{{NS}}}calcPr")
        if calc_pr is None:
            calc_pr = ET.SubElement(workbook_xml, f"{{{NS}}}calcPr")

        before = dict(calc_pr.attrib)

        # Safe performance settings:
        # - keep automatic calculation
        # - stop forcing a full recalculation of the entire workbook
        calc_pr.set("calcMode", "auto")
        calc_pr.set("calcCompleted", "1")
        calc_pr.set("forceFullCalc", "0")
        calc_pr.set("fullCalcOnLoad", "0")

        after = dict(calc_pr.attrib)

        with tempfile.NamedTemporaryFile(delete=False, suffix=WORKBOOK_PATH.suffix) as tmp:
            tmp_path = Path(tmp.name)

        try:
            with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == "xl/workbook.xml":
                        data = ET.tostring(workbook_xml, encoding="utf-8", xml_declaration=False)
                    zout.writestr(item, data)

            shutil.move(tmp_path, WORKBOOK_PATH)
        finally:
            if tmp_path.exists():
                tmp_path.unlink(missing_ok=True)

    print(f"calcPr before: {before}")
    print(f"calcPr after: {after}")
    print("Safe performance optimization applied.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
