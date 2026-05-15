from __future__ import annotations

import json, re, shutil
from datetime import datetime
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET

ROOT = Path(r"C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX")
OFFICIAL = Path(r"C:\Users\VNO024\OneDrive - Maersk Group\Inland Execution Brazil - OPC - OPC (also BPS)\2026 - Info\DSU\Base_DSU2026 - TbM - WK19.xlsm")
FIXED = ROOT / "analysis" / "Base_DSU2026_WK19_rules_hypercare_fix_20260508_173020.xlsm"
TS = datetime.now().strftime("%Y%m%d_%H%M%S")
CANDIDATE = ROOT / "analysis" / f"Base_DSU2026_WK19_conservative_ready_{TS}.xlsm"
REPORT = ROOT / "analysis" / f"Base_DSU2026_WK19_conservative_ready_{TS}_prepare_report.json"
NS_MAIN='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS_REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_PKG_REL='http://schemas.openxmlformats.org/package/2006/relationships'

def norm_target(t):
    t=t.replace('\\','/')
    if t.startswith('/'): t=t.lstrip('/')
    elif not t.startswith('xl/'): t='xl/'+t
    return t

def sheet_map(z):
    wb=ET.fromstring(z.read('xl/workbook.xml'))
    rels=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rm={r.attrib['Id']: norm_target(r.attrib['Target']) for r in rels.findall(f'{{{NS_PKG_REL}}}Relationship')}
    return {s.attrib['name']: rm[s.attrib[f'{{{NS_REL}}}id']] for s in wb.findall(f'.//{{{NS_MAIN}}}sheet')}

def extract_cols(x):
    m=re.search(r'<cols\b.*?</cols>', x, re.S)
    return m.group(0) if m else None

def replace_cols(fx, oc):
    if oc is None:
        return re.sub(r'<cols\b.*?</cols>', '', fx, count=1, flags=re.S)
    if re.search(r'<cols\b.*?</cols>', fx, re.S):
        return re.sub(r'<cols\b.*?</cols>', oc, fx, count=1, flags=re.S)
    m=re.search(r'</sheetFormatPr>', fx)
    if m:
        return fx[:m.end()]+oc+fx[m.end():]
    m=re.search(r'<sheetData\b', fx)
    if m:
        return fx[:m.start()]+oc+fx[m.start():]
    return fx

def set_attr(tag_text, key, val):
    close='/>' if tag_text.rstrip().endswith('/>') else '>'
    body=tag_text[:-2] if close=='/>' else tag_text[:-1]
    if re.search(rf'\b{re.escape(key)}="[^"]*"', body):
        body=re.sub(rf'\b{re.escape(key)}="[^"]*"', f'{key}="{val}"', body)
    else:
        body += f' {key}="{val}"'
    return body+close

def set_calc_flags(wb_xml):
    m=re.search(r'<calcPr\b[^>]*>', wb_xml, re.S)
    if not m:
        calc='<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1" calcOnSave="1"/>'
        return wb_xml.replace('</workbook>', calc+'</workbook>', 1), True
    tag=m.group(0)
    for k,v in [('calcMode','auto'),('fullCalcOnLoad','1'),('forceFullCalc','1'),('calcOnSave','1')]:
        tag=set_attr(tag,k,v)
    return wb_xml[:m.start()]+tag+wb_xml[m.end():], True

shutil.copy2(FIXED, CANDIDATE)
changes={'candidate':str(CANDIDATE),'source':str(FIXED),'official':str(OFFICIAL),'created_at':datetime.now().isoformat(timespec='seconds'),'vba_copied':False,'vba_crc_matches_official':False,'cols_restored':[],'drawings_copied':[],'calc_flags_set':False}
with ZipFile(OFFICIAL,'r') as zo, ZipFile(CANDIDATE,'r') as zc:
    off_map=sheet_map(zo); cand_map=sheet_map(zc)
    repl={}
    if 'xl/vbaProject.bin' in zo.namelist() and 'xl/vbaProject.bin' in zc.namelist():
        repl['xl/vbaProject.bin']=zo.read('xl/vbaProject.bin'); changes['vba_copied']=True
    for sh, op in off_map.items():
        cp=cand_map.get(sh)
        if not cp: continue
        ox=zo.read(op).decode('utf-8'); fx=zc.read(cp).decode('utf-8')
        nx=replace_cols(fx, extract_cols(ox))
        if nx!=fx:
            repl[cp]=nx.encode('utf-8'); changes['cols_restored'].append(sh)
    for d in sorted(n for n in zo.namelist() if re.match(r'xl/drawings/drawing\d+\.xml$', n)):
        if d in zc.namelist():
            repl[d]=zo.read(d); changes['drawings_copied'].append(d)
    wb=zc.read('xl/workbook.xml').decode('utf-8')
    nwb, ch=set_calc_flags(wb)
    if ch:
        repl['xl/workbook.xml']=nwb.encode('utf-8'); changes['calc_flags_set']=True

tmp=CANDIDATE.with_suffix('.tmp.xlsm')
with ZipFile(CANDIDATE,'r') as zin, ZipFile(tmp,'w',compression=ZIP_DEFLATED,allowZip64=True) as zout:
    for item in zin.infolist():
        data=repl.get(item.filename)
        if data is None: data=zin.read(item.filename)
        zout.writestr(item,data)
tmp.replace(CANDIDATE)
with ZipFile(OFFICIAL,'r') as zo, ZipFile(CANDIDATE,'r') as zc:
    changes['zip_test_bad']=zc.testzip()
    changes['sheet_count']=len([n for n in zc.namelist() if n.startswith('xl/worksheets/sheet') and n.endswith('.xml')])
    changes['pivot_tables']=len([n for n in zc.namelist() if n.startswith('xl/pivotTables/pivotTable') and n.endswith('.xml')])
    changes['tables']=len([n for n in zc.namelist() if n.startswith('xl/tables/') and n.endswith('.xml')])
    changes['charts']=len([n for n in zc.namelist() if n.startswith('xl/charts/') and n.endswith('.xml')])
    changes['drawings']=len([n for n in zc.namelist() if n.startswith('xl/drawings/drawing') and n.endswith('.xml')])
    changes['vba_crc_official']=zo.getinfo('xl/vbaProject.bin').CRC
    changes['vba_crc_candidate']=zc.getinfo('xl/vbaProject.bin').CRC
    changes['vba_crc_matches_official']=changes['vba_crc_official']==changes['vba_crc_candidate']
    changes['size']=CANDIDATE.stat().st_size
REPORT.write_text(json.dumps(changes,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(changes,ensure_ascii=False,indent=2))
print('REPORT='+str(REPORT))
