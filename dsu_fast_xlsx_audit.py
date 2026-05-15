from __future__ import annotations
import json, re, sys, time, posixpath
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET
from collections import defaultdict, Counter

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

WORKBOOK = Path(r"C:\Users\VNO024\OneDrive - Maersk Group\2025\Py\DSU 2025\Base_DSU2026.xlsm")
OUT_JSON = Path(r"C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX\dsu_fast_xlsx_audit_result.json")

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKGREL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_DRAW = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

DASHBOARD_SHEETS = {
    "Volume_DS", "Volume_MAO", "Volume_Graph", "Week_Overview",
    "Top_Offenders_Customers", "Top_Offenders_Vendors", "Menu",
    "LastMonth_Overview", "Month2date"
}
STORE_SHEETS = DASHBOARD_SHEETS | {"ROE_wk", "Amb+RoFo+Cap", "Weekly", "Din", "Capacity Plan"}
ERROR_VALUES = {"#N/A", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#NULL!"}


def q(ns, tag):
    return f"{{{ns}}}{tag}"


def localname(tag):
    return tag.split('}', 1)[-1]


def col_to_num(col):
    n = 0
    for ch in col.upper():
        if 'A' <= ch <= 'Z':
            n = n * 26 + ord(ch) - 64
    return n


def num_to_col(n):
    s = ""
    while n:
        n, r = divmod(n-1, 26)
        s = chr(65+r) + s
    return s


def split_addr(addr):
    m = re.match(r"\$?([A-Z]{1,3})\$?(\d+)$", addr)
    if not m:
        return None
    return m.group(1), int(m.group(2))


def range_addrs(ref, max_cells=5000):
    ref = ref.replace('$', '')
    if ':' not in ref:
        p = split_addr(ref)
        return [ref] if p else []
    a, b = ref.split(':', 1)
    pa, pb = split_addr(a), split_addr(b)
    if not pa or not pb:
        return []
    c1, r1 = pa; c2, r2 = pb
    n1, n2 = col_to_num(c1), col_to_num(c2)
    out = []
    for r in range(min(r1, r2), max(r1, r2)+1):
        for c in range(min(n1, n2), max(n1, n2)+1):
            out.append(f"{num_to_col(c)}{r}")
            if len(out) >= max_cells:
                return out
    return out


def parse_dimension(ref):
    if not ref:
        return None
    ref = ref.replace('$','')
    if ':' not in ref:
        a = b = ref
    else:
        a, b = ref.split(':', 1)
    pa, pb = split_addr(a), split_addr(b)
    if not pa or not pb:
        return None
    c1, r1 = pa; c2, r2 = pb
    return {"ref": ref, "start_row": r1, "end_row": r2, "start_col": c1, "end_col": c2, "rows": abs(r2-r1)+1, "cols": abs(col_to_num(c2)-col_to_num(c1))+1}


def load_shared_strings(z):
    if 'xl/sharedStrings.xml' not in z.namelist():
        return []
    root = ET.fromstring(z.read('xl/sharedStrings.xml'))
    out = []
    for si in root.findall(q(NS_MAIN, 'si')):
        texts = []
        for t in si.iter(q(NS_MAIN, 't')):
            texts.append(t.text or '')
        out.append(''.join(texts))
    return out


def cell_value(c, shared):
    t = c.attrib.get('t')
    v = c.find(q(NS_MAIN, 'v'))
    raw = v.text if v is not None else None
    if t == 's':
        try:
            return shared[int(raw)] if raw is not None else None
        except Exception:
            return raw
    if t == 'inlineStr':
        texts = []
        for tt in c.iter(q(NS_MAIN, 't')):
            texts.append(tt.text or '')
        return ''.join(texts)
    if t == 'e':
        return raw
    if t in ('str', 'b'):
        return raw
    if raw is None:
        return None
    try:
        f = float(raw)
        return int(f) if f.is_integer() else f
    except Exception:
        return raw


def rels_for(z, source_path):
    # source_path like xl/worksheets/sheet1.xml -> xl/worksheets/_rels/sheet1.xml.rels
    folder, name = posixpath.split(source_path)
    rel_path = posixpath.join(folder, '_rels', name + '.rels')
    rels = {}
    if rel_path not in z.namelist():
        return rels
    root = ET.fromstring(z.read(rel_path))
    for rel in root.findall(q(NS_PKGREL, 'Relationship')):
        rid = rel.attrib.get('Id')
        target = rel.attrib.get('Target')
        typ = rel.attrib.get('Type')
        if target and not target.startswith('/'):
            target = posixpath.normpath(posixpath.join(folder, target))
        elif target:
            target = target.lstrip('/')
        rels[rid] = {'target': target, 'type': typ}
    return rels


def workbook_maps(z):
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rel_root = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    wb_rels = {}
    for rel in rel_root.findall(q(NS_PKGREL, 'Relationship')):
        target = rel.attrib['Target']
        if not target.startswith('/'):
            target = posixpath.normpath(posixpath.join('xl', target))
        else:
            target = target.lstrip('/')
        wb_rels[rel.attrib['Id']] = {'target': target, 'type': rel.attrib.get('Type')}
    sheets = []
    for sh in wb.find(q(NS_MAIN, 'sheets')).findall(q(NS_MAIN, 'sheet')):
        rid = sh.attrib.get(q(NS_REL, 'id'))
        sheets.append({'name': sh.attrib['name'], 'sheetId': sh.attrib.get('sheetId'), 'rid': rid, 'path': wb_rels.get(rid, {}).get('target')})
    pivot_caches = {}
    pc_parent = wb.find(q(NS_MAIN, 'pivotCaches'))
    if pc_parent is not None:
        for pc in pc_parent.findall(q(NS_MAIN, 'pivotCache')):
            rid = pc.attrib.get(q(NS_REL, 'id'))
            pivot_caches[pc.attrib.get('cacheId')] = wb_rels.get(rid, {}).get('target')
    return sheets, wb_rels, pivot_caches


def parse_sheet(z, sheet, shared, store=False, dashboard=False):
    path = sheet['path']
    stats = {
        'sheet': sheet['name'], 'path': path, 'dimension': None, 'formula_count': 0,
        'formula_error_count': 0, 'formula_error_samples': [], 'non_formula_error_count': 0,
        'blank_formula_result_count': None, 'blank_formula_samples': [], 'stored_cell_count': 0,
    }
    cells = {} if store else None
    formulas = {} if store else None
    if path not in z.namelist():
        stats['missing_xml'] = True
        return stats, cells, formulas
    root = ET.fromstring(z.read(path))
    dim = root.find(q(NS_MAIN, 'dimension'))
    stats['dimension'] = parse_dimension(dim.attrib.get('ref') if dim is not None else None)
    if dashboard:
        stats['blank_formula_result_count'] = 0
    for c in root.iter(q(NS_MAIN, 'c')):
        addr = c.attrib.get('r')
        if not addr:
            continue
        f = c.find(q(NS_MAIN, 'f'))
        has_formula = f is not None
        val = cell_value(c, shared)
        if has_formula:
            stats['formula_count'] += 1
            if c.attrib.get('t') == 'e' or (isinstance(val, str) and val in ERROR_VALUES):
                stats['formula_error_count'] += 1
                if len(stats['formula_error_samples']) < 20:
                    stats['formula_error_samples'].append({'address': addr, 'value': val, 'formula': (f.text or '')[:220]})
            if dashboard and (val is None or val == ''):
                stats['blank_formula_result_count'] += 1
                if len(stats['blank_formula_samples']) < 12:
                    stats['blank_formula_samples'].append({'address': addr, 'formula': (f.text or '')[:220]})
        else:
            if c.attrib.get('t') == 'e' or (isinstance(val, str) and val in ERROR_VALUES):
                stats['non_formula_error_count'] += 1
        if store:
            cells[addr.replace('$','')] = val
            if has_formula:
                formulas[addr.replace('$','')] = f.text or ''
    if store:
        stats['stored_cell_count'] = len(cells)
    return stats, cells, formulas


def parse_tables(z):
    tables = {}
    for path in sorted(n for n in z.namelist() if n.startswith('xl/tables/table') and n.endswith('.xml')):
        root = ET.fromstring(z.read(path))
        name = root.attrib.get('name') or root.attrib.get('displayName')
        ref = root.attrib.get('ref')
        dim = parse_dimension(ref)
        tables[name] = {'path': path, 'displayName': root.attrib.get('displayName'), 'ref': ref, 'dimension': dim, 'data_rows': (dim['rows']-1 if dim else None)}
    return tables


def parse_cache_def(z, path):
    if not path or path not in z.namelist():
        return {'path': path, 'missing': True}
    root = ET.fromstring(z.read(path))
    src = root.find(q(NS_MAIN, 'cacheSource'))
    ws_src = src.find(q(NS_MAIN, 'worksheetSource')) if src is not None else None
    return {
        'path': path,
        'recordCount': root.attrib.get('recordCount'),
        'saveData': root.attrib.get('saveData'),
        'refreshOnLoad': root.attrib.get('refreshOnLoad'),
        'refreshedBy': root.attrib.get('refreshedBy'),
        'refreshedDate': root.attrib.get('refreshedDate'),
        'source': dict(ws_src.attrib) if ws_src is not None else None,
    }


def range_values_summary(sheet_cells, sheet_name, ref):
    addrs = range_addrs(ref, max_cells=10000)
    vals = [sheet_cells.get(sheet_name, {}).get(a) for a in addrs]
    blanks = sum(1 for v in vals if v is None or v == '')
    errors = sum(1 for v in vals if isinstance(v, str) and v in ERROR_VALUES)
    nums = sum(1 for v in vals if isinstance(v, (int, float)))
    return {'ref': ref, 'cell_count': len(addrs), 'blank_count': blanks, 'error_count': errors, 'number_count': nums, 'sample_values': [v for v in vals[:8]]}


def parse_pivots(z, sheets, pivot_cache_map, sheet_cells):
    caches = {cid: parse_cache_def(z, path) for cid, path in pivot_cache_map.items()}
    pivots = []
    for sheet in sheets:
        rels = rels_for(z, sheet['path'])
        for rel in rels.values():
            if rel.get('type', '').endswith('/pivotTable'):
                ppath = rel['target']
                if ppath not in z.namelist():
                    pivots.append({'sheet': sheet['name'], 'path': ppath, 'missing': True})
                    continue
                root = ET.fromstring(z.read(ppath))
                loc = root.find(q(NS_MAIN, 'location'))
                cache_id = root.attrib.get('cacheId')
                data_fields = []
                dfs = root.find(q(NS_MAIN, 'dataFields'))
                if dfs is not None:
                    for df in dfs.findall(q(NS_MAIN, 'dataField')):
                        data_fields.append(dict(df.attrib))
                page_fields = []
                pfs = root.find(q(NS_MAIN, 'pageFields'))
                if pfs is not None:
                    for pf in pfs.findall(q(NS_MAIN, 'pageField')):
                        page_fields.append(dict(pf.attrib))
                ref = loc.attrib.get('ref') if loc is not None else None
                dim = parse_dimension(ref)
                body_summary = None
                if dim and sheet['name'] in sheet_cells:
                    vals = []
                    for a in range_addrs(ref, max_cells=20000):
                        vals.append(sheet_cells[sheet['name']].get(a))
                    body_summary = {
                        'cell_count_scanned': len(vals),
                        'blank_count': sum(1 for v in vals if v is None or v == ''),
                        'error_count': sum(1 for v in vals if isinstance(v, str) and v in ERROR_VALUES),
                        'nonblank_count': sum(1 for v in vals if v not in (None,'')),
                    }
                cache = caches.get(cache_id, {})
                # source sanity check for tables
                pivots.append({
                    'sheet': sheet['name'], 'name': root.attrib.get('name'), 'path': ppath, 'cacheId': cache_id,
                    'ref': ref, 'dimension': dim, 'dataOnRows': root.attrib.get('dataOnRows'),
                    'data_fields': data_fields, 'page_fields': page_fields, 'cache': cache,
                    'body_summary': body_summary,
                })
    return pivots


def parse_charts(z, sheets, sheet_cells):
    charts = []
    for sheet in sheets:
        rels = rels_for(z, sheet['path'])
        drawing_targets = [r['target'] for r in rels.values() if r.get('type','').endswith('/drawing')]
        for dpath in drawing_targets:
            if dpath not in z.namelist():
                continue
            drels = rels_for(z, dpath)
            root = ET.fromstring(z.read(dpath))
            chart_rids = []
            for elem in root.iter():
                if localname(elem.tag) == 'chart':
                    rid = elem.attrib.get(q(NS_REL, 'id'))
                    if rid:
                        chart_rids.append(rid)
            for rid in chart_rids:
                cpath = drels.get(rid, {}).get('target')
                if not cpath or cpath not in z.namelist():
                    charts.append({'sheet': sheet['name'], 'chart_path': cpath, 'missing': True})
                    continue
                croot = ET.fromstring(z.read(cpath))
                series = []
                for ser in croot.iter(q(NS_CHART, 'ser')):
                    formulas = []
                    for f in ser.iter(q(NS_CHART, 'f')):
                        if f.text:
                            formulas.append(f.text)
                    caches = []
                    for cache_tag in ['numCache', 'strCache']:
                        for cache in ser.iter(q(NS_CHART, cache_tag)):
                            pts = []
                            for pt in cache.findall(q(NS_CHART, 'pt')):
                                v = pt.find(q(NS_CHART, 'v'))
                                pts.append(v.text if v is not None else None)
                            caches.append({'type': cache_tag, 'count': len(pts), 'blank_count': sum(1 for v in pts if v in (None,'')), 'error_count': sum(1 for v in pts if isinstance(v,str) and v in ERROR_VALUES), 'sample': pts[:8]})
                    source_summaries = []
                    for formula in formulas:
                        # Supports 'Sheet Name'!$A$1:$B$2 and Sheet!$A$1:$B$2
                        for m in re.finditer(r"(?:'([^']+)'|([A-Za-z0-9_ ]+))!(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)", formula):
                            sname = m.group(1) or m.group(2)
                            ref = m.group(3)
                            if sname in sheet_cells:
                                source_summaries.append(range_values_summary(sheet_cells, sname, ref))
                            else:
                                source_summaries.append({'sheet': sname, 'ref': ref, 'not_evaluated': True})
                    series.append({'formulas': formulas, 'caches': caches, 'source_summaries': source_summaries})
                charts.append({'sheet': sheet['name'], 'drawing': dpath, 'chart_path': cpath, 'series_count': len(series), 'series': series})
    return charts


def sv(v):
    if v is None or v == '': return ''
    if isinstance(v, float) and v.is_integer(): return str(int(v))
    return str(v)


def count_records(records, criteria):
    c = 0
    for r in records:
        ok = True
        for field, op, value in criteria:
            rv = sv(r.get(field)); vv = sv(value)
            if op == 'eq' and rv != vv: ok = False; break
            if op == 'ne' and rv == vv: ok = False; break
        if ok: c += 1
    return c


def make_roe_records(cells, table_dim):
    # Use fixed source columns from ROE_wk table family.
    columns = {
        'Centro de Custo': 'AI', 'Weeknum': 'AR', 'Volume': 'AV', 'Porto': 'AY',
        'OTD ajustado': 'BB', 'Atrasado?': 'BD', 'OTO Out': 'BL', 'Especiais': 'BO',
        'SLA Ag': 'AT', 'Justificativa': 'BC', 'Provedor': 'D', 'Cliente Proposta': 'M',
        'AtrasoRev': 'BP', '% OTO Client': 'BJ', '% OTO Provider': 'BI'
    }
    end = table_dim['end_row'] if table_dim else 9205
    records = []
    for r in range(2, end+1):
        rec = {k: cells.get(f'{col}{r}') for k, col in columns.items()}
        if any(v not in (None,'') for v in rec.values()):
            records.append(rec)
    return records


def moves(records, port=None, center=None):
    crit = [('Volume','eq','Ok'), ('Weeknum','eq',20)]
    if port: crit.append(('Porto','eq',port))
    if center == 'CAB': crit.append(('Centro de Custo','eq','Aliança'))
    elif center == 'DS': crit.append(('Centro de Custo','ne','Aliança'))
    return count_records(records, crit)


def oto(records, port=None, center=None):
    den = [('Volume','eq','Ok'), ('Weeknum','eq',20), ('OTD ajustado','ne','Sem Preenchimento'), ('OTO Out','eq','N')]
    num = [('Volume','eq','Ok'), ('Weeknum','eq',20), ('OTD ajustado','eq','Atrasado'), ('OTO Out','eq','N'), ('Especiais','ne','Especial')]
    if port:
        den.append(('Porto','eq',port)); num.append(('Porto','eq',port))
    if center == 'CAB':
        den.append(('Centro de Custo','eq','Aliança')); num.append(('Centro de Custo','eq','Aliança'))
    elif center == 'DS':
        den.append(('Centro de Custo','ne','Aliança')); num.append(('Centro de Custo','ne','Aliança'))
    d = count_records(records, den); n = count_records(records, num)
    return {'denominator': d, 'delay_penalty': n, 'value': None if d == 0 else 1 - n/d}


def close(a,b,tol=1e-6):
    if a in (None,'') and b is None: return True
    if a in (None,'') or b is None: return False
    try: return abs(float(a)-float(b)) <= tol
    except Exception: return sv(a)==sv(b)


def validate_week_overview(sheet_cells, records):
    cells = sheet_cells['Week_Overview']
    rows = [2,3,6,7,8,11,12,13,16,17,18,19,20,23]
    out=[]
    for r in rows:
        label = cells.get(f'A{r}')
        port = None if r == 23 else cells.get(f'AG{r}')
        calc = {
            'row': r, 'label': label, 'port': port,
            'moves_cab': moves(records, port, 'CAB'),
            'moves_ds': moves(records, port, 'DS'),
            'moves_total': moves(records, port, None),
            'oto_cab': oto(records, port, 'CAB'),
            'oto_ds': oto(records, port, 'DS'),
            'oto_total': oto(records, port, None),
            'sheet': {
                'D': cells.get(f'D{r}'), 'H': cells.get(f'H{r}'), 'K': cells.get(f'K{r}'),
                'N': cells.get(f'N{r}'), 'P': cells.get(f'P{r}'), 'Q': cells.get(f'Q{r}')
            }
        }
        diffs=[]
        if not close(calc['sheet']['D'], calc['moves_cab']): diffs.append('D moves CAB')
        if not close(calc['sheet']['H'], calc['moves_ds']): diffs.append('H moves DS')
        if not close(calc['sheet']['K'], calc['moves_total']): diffs.append('K moves total')
        tol_np = 0.006 if r == 23 else 1e-6
        if not close(calc['sheet']['N'], calc['oto_cab']['value'], tol_np): diffs.append('N OTO CAB')
        if not close(calc['sheet']['P'], calc['oto_ds']['value'], tol_np): diffs.append('P OTO DS')
        if not close(calc['sheet']['Q'], calc['oto_total']['value']): diffs.append('Q OTO Total')
        calc['diffs']=diffs
        out.append(calc)
    # Region rows sum check from visible port rows.
    groups = {'North': ([2,3],4), 'Northeast': ([6,7,8],9), 'Southeast': ([11,12,13],14), 'South': ([16,17,18,19,20],21)}
    region=[]
    for name,(children,row) in groups.items():
        diffs=[]
        for col in ['D','H','K']:
            expected=sum((cells.get(f'{col}{cr}') or 0) for cr in children)
            if not close(cells.get(f'{col}{row}'), expected): diffs.append(f'{col} sum')
        region.append({'region': name, 'row': row, 'diffs': diffs})
    return out, region


def validate_graph_sheets(sheet_cells, records):
    wo = sheet_cells['Week_Overview']
    out=[]
    # Expected MAO row 2.
    expected = {'moves_total': wo.get('K2'), 'moves_cab': wo.get('D2'), 'moves_ds': wo.get('H2'), 'oto_cab': wo.get('N2'), 'oto_ds': wo.get('P2'), 'oto_total': wo.get('Q2')}
    backlog = count_records(records, [('Volume','eq','Ok'),('Weeknum','eq',20),('Porto','eq','Manaus'),('OTO Out','eq','N')])
    for sheet in ['Volume_Graph','Volume_DS','Volume_MAO']:
        c = sheet_cells.get(sheet, {})
        checks=[]
        checks.append({'check':'M2 week', 'sheet': c.get('M2'), 'expected':20, 'ok': close(c.get('M2'),20)})
        checks.append({'check':'M3 port', 'sheet': c.get('M3'), 'expected':'Manaus', 'ok': sv(c.get('M3'))=='Manaus'})
        checks.append({'check':'P1 backlog', 'sheet': c.get('P1'), 'expected':backlog, 'ok': close(c.get('P1'),backlog)})
        checks.append({'check':'M9 total volume vs Week_Overview K2', 'sheet': c.get('M9'), 'expected':expected['moves_total'], 'ok': close(c.get('M9'),expected['moves_total'])})
        if sheet == 'Volume_MAO':
            cab = (c.get('M12') or 0) + (c.get('M13') or 0)
            checks.append({'check':'M12+M13 CAB split vs Week_Overview D2', 'sheet': cab, 'expected':expected['moves_cab'], 'ok': close(cab, expected['moves_cab'])})
            checks.append({'check':'M14 DS vs Week_Overview H2', 'sheet': c.get('M14'), 'expected':expected['moves_ds'], 'ok': close(c.get('M14'),expected['moves_ds'])})
            checks.append({'check':'M19 OTO CAB vs Week_Overview N2', 'sheet': c.get('M19'), 'expected':expected['oto_cab'], 'ok': close(c.get('M19'),expected['oto_cab'],0.006)})
        elif sheet == 'Volume_DS':
            checks.append({'check':'M11 CAB vs Week_Overview D2', 'sheet': c.get('M11'), 'expected':expected['moves_cab'], 'ok': close(c.get('M11'),expected['moves_cab'])})
            checks.append({'check':'M12 DS vs Week_Overview H2', 'sheet': c.get('M12'), 'expected':expected['moves_ds'], 'ok': close(c.get('M12'),expected['moves_ds'])})
            checks.append({'check':'M16 OTO DS vs Week_Overview P2', 'sheet': c.get('M16'), 'expected':expected['oto_ds'], 'ok': close(c.get('M16'),expected['oto_ds'],0.006)})
        else:
            checks.append({'check':'M12 CAB vs Week_Overview D2', 'sheet': c.get('M12'), 'expected':expected['moves_cab'], 'ok': close(c.get('M12'),expected['moves_cab'])})
            checks.append({'check':'M13 DS vs Week_Overview H2', 'sheet': c.get('M13'), 'expected':expected['moves_ds'], 'ok': close(c.get('M13'),expected['moves_ds'])})
            checks.append({'check':'M17 OTO CAB vs Week_Overview N2', 'sheet': c.get('M17'), 'expected':expected['oto_cab'], 'ok': close(c.get('M17'),expected['oto_cab'],0.006)})
        out.append({'sheet':sheet, 'checks':checks, 'failed': [x for x in checks if not x['ok']]})
    return out


def main():
    result={'workbook':str(WORKBOOK), 'started_at':time.strftime('%Y-%m-%d %H:%M:%S')}
    with ZipFile(WORKBOOK) as z:
        bad=z.testzip()
        result['zip_test']='OK' if bad is None else bad
        result['has_vba']='xl/vbaProject.bin' in z.namelist()
        shared=load_shared_strings(z)
        sheets, wb_rels, pivot_cache_map=workbook_maps(z)
        result['sheet_order']=[s['name'] for s in sheets]
        result['tables']=parse_tables(z)
        # Store cached cells for all dashboard sheets plus any sheet that owns pivots/charts,
        # so pivot outputs and chart source ranges can be checked without Excel COM.
        dynamic_store_sheets=set(STORE_SHEETS)
        for sh_probe in sheets:
            rels_probe=rels_for(z, sh_probe['path'])
            if any(r.get('type','').endswith('/pivotTable') or r.get('type','').endswith('/drawing') for r in rels_probe.values()):
                dynamic_store_sheets.add(sh_probe['name'])
        result['stored_sheets']=sorted(dynamic_store_sheets)
        sheet_stats=[]; sheet_cells={}; sheet_formulas={}
        for sh in sheets:
            store=sh['name'] in dynamic_store_sheets
            stats,cells,formulas=parse_sheet(z, sh, shared, store=store, dashboard=sh['name'] in DASHBOARD_SHEETS)
            sheet_stats.append(stats)
            if store:
                sheet_cells[sh['name']]=cells
                sheet_formulas[sh['name']]=formulas
        result['sheets']=sheet_stats
        result['pivots']=parse_pivots(z, sheets, pivot_cache_map, sheet_cells)
        result['charts']=parse_charts(z, sheets, sheet_cells)
    # Data validations from cached workbook values.
    roe_table=result['tables'].get('ROE_wk')
    records=make_roe_records(sheet_cells['ROE_wk'], roe_table['dimension'] if roe_table else None)
    result['roe_summary']={
        'records_with_any_relevant_value':len(records),
        'ROE_wk_table_rows': roe_table.get('data_rows') if roe_table else None,
        'week20_volume_ok': count_records(records,[('Weeknum','eq',20),('Volume','eq','Ok')]),
        'week20_volume_ok_cab': count_records(records,[('Weeknum','eq',20),('Volume','eq','Ok'),('Centro de Custo','eq','Aliança')]),
        'week20_volume_ok_ds': count_records(records,[('Weeknum','eq',20),('Volume','eq','Ok'),('Centro de Custo','ne','Aliança')]),
        'week20_sem_preenchimento': count_records(records,[('Weeknum','eq',20),('Volume','eq','Ok'),('OTD ajustado','eq','Sem Preenchimento')]),
        'week20_oto_denominator_total': oto(records,None,None)['denominator'],
        'week20_oto_delay_penalty_total': oto(records,None,None)['delay_penalty'],
    }
    wk, regions=validate_week_overview(sheet_cells, records)
    result['week_overview_validation']=wk
    result['week_overview_region_sum_validation']=regions
    result['graph_sheet_validation']=validate_graph_sheets(sheet_cells, records)
    # Summarize issues.
    issues=[]
    for st in sheet_stats:
        if st['formula_error_count']:
            issues.append({'severity':'attention','type':'formula_errors','sheet':st['sheet'],'count':st['formula_error_count'],'samples':st['formula_error_samples'][:10]})
    for p in result['pivots']:
        bs=p.get('body_summary') or {}
        if p.get('missing') or not p.get('ref') or (bs and bs.get('nonblank_count',1)==0) or (bs and bs.get('error_count',0)>0):
            issues.append({'severity':'warning','type':'pivot_output_issue','sheet':p.get('sheet'),'pivot':p.get('name'),'ref':p.get('ref'),'body_summary':bs})
        cache=p.get('cache') or {}
        source=cache.get('source') or {}
        if source.get('name') == 'ROE_wk' and str(cache.get('recordCount')) != str(result['tables'].get('ROE_wk',{}).get('data_rows')):
            issues.append({'severity':'attention','type':'pivot_cache_record_count_mismatch','sheet':p.get('sheet'),'pivot':p.get('name'),'cache_records':cache.get('recordCount'),'roe_table_rows':result['tables'].get('ROE_wk',{}).get('data_rows')})
    for ch in result['charts']:
        if ch.get('series_count',0)==0:
            issues.append({'severity':'attention','type':'chart_no_series','sheet':ch.get('sheet'),'chart_path':ch.get('chart_path')})
        for idx,ser in enumerate(ch.get('series',[]),start=1):
            if any(cache.get('error_count',0)>0 for cache in ser.get('caches',[])):
                issues.append({'severity':'attention','type':'chart_cache_error_values','sheet':ch.get('sheet'),'chart_path':ch.get('chart_path'),'series':idx})
            for ss in ser.get('source_summaries',[]):
                if ss.get('error_count',0)>0:
                    issues.append({'severity':'attention','type':'chart_source_has_error_values','sheet':ch.get('sheet'),'chart_path':ch.get('chart_path'),'series':idx,'source':ss})
                if ss.get('cell_count',0)>0 and ss.get('blank_count')==ss.get('cell_count'):
                    issues.append({'severity':'attention','type':'chart_source_all_blank','sheet':ch.get('sheet'),'chart_path':ch.get('chart_path'),'series':idx,'source':ss})
    for row in wk:
        if row['diffs']:
            issues.append({'severity':'warning','type':'week_overview_mismatch','row':row['row'],'label':row['label'],'diffs':row['diffs']})
    for reg in regions:
        if reg['diffs']:
            issues.append({'severity':'warning','type':'week_overview_region_sum_mismatch',**reg})
    for gv in result['graph_sheet_validation']:
        if gv['failed']:
            issues.append({'severity':'warning','type':'graph_sheet_mismatch','sheet':gv['sheet'],'failed':gv['failed']})
    result['issues']=issues
    result['finished_at']=time.strftime('%Y-%m-%d %H:%M:%S')
    OUT_JSON.write_text(json.dumps(result,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
    return result

if __name__=='__main__':
    res=main()
    print('WROTE', OUT_JSON)
    print(json.dumps({
        'zip_test':res['zip_test'], 'has_vba':res['has_vba'], 'sheet_count':len(res['sheets']),
        'pivot_count':len(res['pivots']), 'chart_count':len(res['charts']), 'issue_count':len(res['issues']),
        'roe_summary':res['roe_summary'],
        'issues_preview':res['issues'][:25]
    },ensure_ascii=False,indent=2,default=str))

