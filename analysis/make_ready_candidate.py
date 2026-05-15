from __future__ import annotations

import json
import re
import shutil
from collections import Counter
from datetime import datetime
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET

ROOT = Path(r"C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX")
OFFICIAL = Path(r"C:\Users\VNO024\OneDrive - Maersk Group\Inland Execution Brazil - OPC - OPC (also BPS)\2026 - Info\DSU\Base_DSU2026 - TbM - WK19.xlsm")
FIXED = ROOT / "analysis" / "Base_DSU2026_WK19_rules_hypercare_fix_20260508_173020.xlsm"
TS = datetime.now().strftime("%Y%m%d_%H%M%S")
CANDIDATE = ROOT / "analysis" / f"Base_DSU2026_WK19_ready_to_apply_{TS}.xlsm"
REPORT = ROOT / "analysis" / f"Base_DSU2026_WK19_ready_to_apply_{TS}_prepare_report.json"

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
ET.register_namespace("", NS_MAIN)
ET.register_namespace("r", NS_REL)


def norm_target(target: str) -> str:
    target = target.replace('\\', '/')
    if target.startswith('/'):
        target = target.lstrip('/')
    elif not target.startswith('xl/'):
        target = 'xl/' + target
    return target


def workbook_sheet_map(z: ZipFile) -> dict[str, str]:
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rel_map = {rel.attrib['Id']: norm_target(rel.attrib['Target']) for rel in rels.findall(f'{{{NS_PKG_REL}}}Relationship')}
    out = {}
    for sheet in wb.findall(f'.//{{{NS_MAIN}}}sheet'):
        name = sheet.attrib['name']
        rid = sheet.attrib[f'{{{NS_REL}}}id']
        out[name] = rel_map[rid]
    return out


def extract_cols(xml_text: str) -> str | None:
    m = re.search(r'<cols\b.*?</cols>', xml_text, flags=re.S)
    return m.group(0) if m else None


def replace_cols(fixed_xml: str, official_cols: str | None) -> tuple[str, bool]:
    has_fixed = re.search(r'<cols\b.*?</cols>', fixed_xml, flags=re.S)
    if official_cols is None:
        if has_fixed:
            return re.sub(r'<cols\b.*?</cols>', '', fixed_xml, count=1, flags=re.S), True
        return fixed_xml, False
    if has_fixed:
        return re.sub(r'<cols\b.*?</cols>', official_cols, fixed_xml, count=1, flags=re.S), True
    # Standard order is sheetPr, dimension, sheetViews, sheetFormatPr, cols, sheetData.
    m = re.search(r'</sheetFormatPr>', fixed_xml)
    if m:
        return fixed_xml[:m.end()] + official_cols + fixed_xml[m.end():], True
    m = re.search(r'<sheetData\b', fixed_xml)
    if m:
        return fixed_xml[:m.start()] + official_cols + fixed_xml[m.start():], True
    return fixed_xml, False


def set_attrs_in_tag(tag_text: str, attrs: dict[str, str]) -> str:
    # Preserve original tag text and simply add/replace attributes.
    close = '/>' if tag_text.rstrip().endswith('/>') else '>'
    body = tag_text[:-2] if close == '/>' else tag_text[:-1]
    for k, v in attrs.items():
        if re.search(rf'\b{re.escape(k)}="[^"]*"', body):
            body = re.sub(rf'\b{re.escape(k)}="[^"]*"', f'{k}="{v}"', body)
        else:
            body += f' {k}="{v}"'
    return body + close


def set_first_tag_attrs(xml_text: str, tag: str, attrs: dict[str, str]) -> tuple[str, bool]:
    pat = rf'<{re.escape(tag)}\b[^>]*>'
    m = re.search(pat, xml_text, flags=re.S)
    if not m:
        return xml_text, False
    return xml_text[:m.start()] + set_attrs_in_tag(m.group(0), attrs) + xml_text[m.end():], True


def set_all_tag_attrs(xml_text: str, tag: str, attrs: dict[str, str]) -> tuple[str, int]:
    pat = rf'<{re.escape(tag)}\b[^>]*>'
    count = 0
    def repl(m):
        nonlocal count
        count += 1
        return set_attrs_in_tag(m.group(0), attrs)
    return re.sub(pat, repl, xml_text, flags=re.S), count


def set_workbook_calc(xml_text: str) -> tuple[str, bool]:
    attrs = {
        'calcMode': 'auto',
        'fullCalcOnLoad': '1',
        'forceFullCalc': '1',
        'calcOnSave': '1',
    }
    new_xml, changed = set_first_tag_attrs(xml_text, 'calcPr', attrs)
    if changed:
        return new_xml, True
    calc = '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1" calcOnSave="1"/>'
    if '</workbook>' in xml_text:
        return xml_text.replace('</workbook>', calc + '</workbook>', 1), True
    return xml_text, False


def crc_of(z: ZipFile, name: str) -> int | None:
    try:
        return z.getinfo(name).CRC
    except KeyError:
        return None


def main() -> int:
    if not OFFICIAL.exists():
        raise FileNotFoundError(OFFICIAL)
    if not FIXED.exists():
        raise FileNotFoundError(FIXED)

    shutil.copy2(FIXED, CANDIDATE)

    stats = {
        'created_at': datetime.now().isoformat(timespec='seconds'),
        'official': str(OFFICIAL),
        'fixed_source': str(FIXED),
        'candidate': str(CANDIDATE),
        'copied_vba_project': False,
        'vba_crc_matches_official': False,
        'sheet_count_official': None,
        'sheet_count_candidate': None,
        'cols_restored_sheets': [],
        'cols_unchanged_sheets': [],
        'drawings_copied': [],
        'workbook_calc_flags_set': False,
        'pivot_cache_defs_touched': 0,
        'pivot_table_defs_touched': 0,
        'connections_touched': 0,
        'content_types_preserved': True,
    }

    with ZipFile(OFFICIAL, 'r') as zoff, ZipFile(CANDIDATE, 'r') as zcand:
        official_sheet_map = workbook_sheet_map(zoff)
        candidate_sheet_map = workbook_sheet_map(zcand)
        stats['sheet_count_official'] = len(official_sheet_map)
        stats['sheet_count_candidate'] = len(candidate_sheet_map)

        replacements: dict[str, bytes] = {}

        # Preserve original VBA project exactly.
        if 'xl/vbaProject.bin' in zoff.namelist() and 'xl/vbaProject.bin' in zcand.namelist():
            replacements['xl/vbaProject.bin'] = zoff.read('xl/vbaProject.bin')
            stats['copied_vba_project'] = True

        # Restore visible column width/hidden settings sheet by sheet.
        for sheet_name, off_path in official_sheet_map.items():
            cand_path = candidate_sheet_map.get(sheet_name)
            if not cand_path or off_path not in zoff.namelist() or cand_path not in zcand.namelist():
                continue
            off_xml = zoff.read(off_path).decode('utf-8')
            cand_xml = zcand.read(cand_path).decode('utf-8')
            new_xml, changed = replace_cols(cand_xml, extract_cols(off_xml))
            if changed and new_xml != cand_xml:
                replacements[cand_path] = new_xml.encode('utf-8')
                stats['cols_restored_sheets'].append(sheet_name)
            else:
                stats['cols_unchanged_sheets'].append(sheet_name)

        # Restore drawing anchors/shape positions from official while preserving chart data parts.
        official_drawings = sorted(n for n in zoff.namelist() if re.match(r'xl/drawings/drawing\d+\.xml$', n))
        cand_names = set(zcand.namelist())
        for drawing in official_drawings:
            if drawing in cand_names:
                replacements[drawing] = zoff.read(drawing)
                stats['drawings_copied'].append(drawing)

        # Force automatic full calculation on open/save.
        wb_xml = zcand.read('xl/workbook.xml').decode('utf-8')
        wb_xml2, changed = set_workbook_calc(wb_xml)
        if changed:
            replacements['xl/workbook.xml'] = wb_xml2.encode('utf-8')
            stats['workbook_calc_flags_set'] = True

        # Pivot cache: refresh on open, enable refresh, do not keep stale missing items.
        for name in sorted(n for n in zcand.namelist() if n.startswith('xl/pivotCache/pivotCacheDefinition') and n.endswith('.xml')):
            xml = zcand.read(name).decode('utf-8')
            xml2, changed = set_first_tag_attrs(xml, 'pivotCacheDefinition', {
                'refreshOnLoad': '1',
                'enableRefresh': '1',
                'saveData': '1',
                'missingItemsLimit': '0',
            })
            if changed:
                replacements[name] = xml2.encode('utf-8')
                stats['pivot_cache_defs_touched'] += 1

        # Pivot tables: preserve formatting/widths where Excel supports it.
        for name in sorted(n for n in zcand.namelist() if n.startswith('xl/pivotTables/pivotTable') and n.endswith('.xml')):
            xml = zcand.read(name).decode('utf-8')
            xml2, changed = set_first_tag_attrs(xml, 'pivotTableDefinition', {
                'preserveFormatting': '1',
                'useAutoFormatting': '0',
                'applyWidthHeightFormats': '0',
            })
            if changed:
                replacements[name] = xml2.encode('utf-8')
                stats['pivot_table_defs_touched'] += 1

        # Workbook connections: refresh on open where Excel supports it.
        if 'xl/connections.xml' in zcand.namelist():
            xml = zcand.read('xl/connections.xml').decode('utf-8')
            xml2, count = set_all_tag_attrs(xml, 'connection', {'refreshOnLoad': '1', 'saveData': '1'})
            if count:
                replacements['xl/connections.xml'] = xml2.encode('utf-8')
                stats['connections_touched'] = count

    # Rewrite package with replacements.
    tmp = CANDIDATE.with_suffix('.tmp.xlsm')
    if tmp.exists():
        tmp.unlink()
    with ZipFile(CANDIDATE, 'r') as zin, ZipFile(tmp, 'w', compression=ZIP_DEFLATED, allowZip64=True) as zout:
        for item in zin.infolist():
            data = replacements.get(item.filename)
            if data is None:
                data = zin.read(item.filename)
            zout.writestr(item, data)
    tmp.replace(CANDIDATE)

    with ZipFile(OFFICIAL, 'r') as zoff, ZipFile(CANDIDATE, 'r') as zcand:
        stats['vba_crc_official'] = crc_of(zoff, 'xl/vbaProject.bin')
        stats['vba_crc_candidate'] = crc_of(zcand, 'xl/vbaProject.bin')
        stats['vba_crc_matches_official'] = stats['vba_crc_official'] == stats['vba_crc_candidate']
        stats['candidate_has_calc_chain'] = 'xl/calcChain.xml' in zcand.namelist()
        stats['candidate_size'] = CANDIDATE.stat().st_size
        stats['candidate_last_write_time'] = datetime.fromtimestamp(CANDIDATE.stat().st_mtime).isoformat(timespec='seconds')

    REPORT.write_text(json.dumps(stats, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(stats, ensure_ascii=False, indent=2))
    print(f"REPORT={REPORT}")
    return 0

if __name__ == '__main__':
    raise SystemExit(main())
