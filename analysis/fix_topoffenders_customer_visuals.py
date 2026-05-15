from __future__ import annotations
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET
import posixpath, re, json, shutil, html
from datetime import datetime

ROOT=Path(r'C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX')
WORKBOOK=Path(r'C:\Users\VNO024\Downloads\Base_DSU2026 - TbM - WK19.xlsm')
BACKUP_DIR=ROOT/'backups'
ANALYSIS_DIR=ROOT/'analysis'
REPORT=ANALYSIS_DIR/'Base_DSU2026_WK19_top_offenders_customer_fix_report.json'
SHEET='Top_Offenders_Customers'
NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS_REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_PKG='http://schemas.openxmlformats.org/package/2006/relationships'
ET.register_namespace('', NS)
ET.register_namespace('r', NS_REL)

def q(tag): return f'{{{NS}}}{tag}'
def resolve(base,t): return posixpath.normpath(posixpath.join(posixpath.dirname(base),t.replace('\\','/'))) if not t.startswith('/') else t.lstrip('/')
def sheet_map(z):
    wb=ET.fromstring(z.read('xl/workbook.xml')); rels=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rm={r.attrib['Id']:resolve('xl/workbook.xml',r.attrib['Target']) for r in rels.findall(f'{{{NS_PKG}}}Relationship')}
    return {sh.attrib['name']:rm[sh.attrib[f'{{{NS_REL}}}id']] for sh in wb.findall(f'.//{{{NS}}}sheet')}
def shared(z):
    out=[]
    if 'xl/sharedStrings.xml' not in z.namelist(): return out
    with z.open('xl/sharedStrings.xml') as f:
        for ev,el in ET.iterparse(f,events=('end',)):
            if el.tag==q('si'):
                out.append(''.join((t.text or '') for t in el.iter(q('t')))); el.clear()
    return out
def cell_value(c, ss):
    typ=c.attrib.get('t'); v=c.find(q('v')); isel=c.find(q('is'))
    if typ=='s' and v is not None and v.text is not None: return ss[int(v.text)]
    if typ=='inlineStr' and isel is not None: return ''.join((t.text or '') for t in isel.iter(q('t')))
    if v is not None and v.text is not None:
        try: return float(v.text)
        except Exception: return v.text
    return None
def col_to_num(col):
    n=0
    for ch in col: n=n*26+ord(ch)-64
    return n
def split_ref(ref):
    m=re.match(r'([A-Z]+)(\d+)',ref)
    return col_to_num(m.group(1)), int(m.group(2))
def get_or_create_row(sheetData, row_map, r):
    if r in row_map: return row_map[r]
    row=ET.Element(q('row'), {'r':str(r)})
    inserted=False
    children=list(sheetData)
    for i,existing in enumerate(children):
        er=int(existing.attrib.get('r','0'))
        if er>r:
            sheetData.insert(i,row); inserted=True; break
    if not inserted: sheetData.append(row)
    row_map[r]=row
    return row
def get_or_create_cell(row, ref, style=None):
    for c in list(row):
        if c.attrib.get('r')==ref: return c
    new=ET.Element(q('c'), {'r':ref})
    if style: new.attrib['s']=str(style)
    col,_=split_ref(ref)
    inserted=False
    for i,c in enumerate(list(row)):
        cref=c.attrib.get('r')
        if cref and split_ref(cref)[0]>col:
            row.insert(i,new); inserted=True; break
    if not inserted: row.append(new)
    return new
def clear_cell(c, keep_style=True):
    style=c.attrib.get('s') if keep_style else None
    ref=c.attrib.get('r')
    c.attrib.clear(); c.attrib['r']=ref
    if style: c.attrib['s']=style
    for child in list(c): c.remove(child)
def set_formula(c, formula, value=None):
    clear_cell(c)
    f=ET.SubElement(c,q('f')); f.text=formula
    if value is not None and value!='':
        v=ET.SubElement(c,q('v'))
        if isinstance(value,float) and value.is_integer(): value=int(value)
        v.text=str(value)
def set_inline(c, text):
    clear_cell(c)
    c.attrib['t']='inlineStr'
    isel=ET.SubElement(c,q('is')); t=ET.SubElement(isel,q('t')); t.text=str(text)
def set_blank(c):
    clear_cell(c)
def numeric(x):
    try:
        if x is None or x=='': return None
        return float(x)
    except Exception: return None
def add_icon_cf(root, ranges):
    # remove old CF; re-add with explicit ranges.
    for cf in list(root.findall(q('conditionalFormatting'))):
        root.remove(cf)
    insert_index=list(root).index(root.find(q('sheetData')))+1
    priority=1
    for rng in ranges:
        cf=ET.Element(q('conditionalFormatting'), {'sqref':rng})
        rule=ET.SubElement(cf,q('cfRule'), {'type':'iconSet','priority':str(priority)})
        icon=ET.SubElement(rule,q('iconSet'))
        ET.SubElement(icon,q('cfvo'), {'type':'percent','val':'0'})
        ET.SubElement(icon,q('cfvo'), {'type':'num','val':'0.8'})
        ET.SubElement(icon,q('cfvo'), {'type':'num','val':'0.92'})
        root.insert(insert_index,cf)
        insert_index+=1; priority+=1

def main():
    BACKUP_DIR.mkdir(exist_ok=True); ANALYSIS_DIR.mkdir(exist_ok=True)
    backup=BACKUP_DIR/f"Base_DSU2026 - TbM - WK19_before_topoffenders_customer_fix_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsm"
    shutil.copy2(WORKBOOK, backup)
    with ZipFile(WORKBOOK,'r') as zin:
        ss=shared(zin); sm=sheet_map(zin); sheet_path=sm[SHEET]
        root=ET.fromstring(zin.read(sheet_path))
        sheetData=root.find(q('sheetData'))
        rows={int(r.attrib['r']):r for r in sheetData.findall(q('row')) if 'r' in r.attrib}
        cells={}
        for row in sheetData.findall(q('row')):
            for c in row.findall(q('c')):
                ref=c.attrib.get('r')
                if ref: cells[ref]=c
        # Snapshot helper values AF:AH for Hyper Care mirror.
        helper={ref:cell_value(c, ss) for ref,c in cells.items() if re.match(r'(AF|AG|AH)\d+$',ref)}
        # Fix G510 total formula/value.
        b=numeric(cell_value(cells.get('B510'),ss)); cval=numeric(cell_value(cells.get('C510'),ss))
        g_value = None if not b else 1 - (cval or 0)/b
        row=get_or_create_row(sheetData, rows, 510); gc=get_or_create_cell(row,'G510', style=cells.get('G509',ET.Element('x')).attrib.get('s') if cells.get('G509') is not None else None)
        set_formula(gc,'IFERROR(-(C510/B510-1),"")', g_value)
        # Hyper Care Norte mirror T:W, mapped from AF:AH row+14.
        updated_hyper=0; blank_hyper=0
        for r in range(7,511):
            row=get_or_create_row(sheetData, rows, r)
            src=r+14
            tval=helper.get(f'AF{src}')
            uval=helper.get(f'AG{src}')
            wval=helper.get(f'AH{src}')
            for col, source_ref, val in [('T',f'AF{src}',tval),('U',f'AG{src}',uval),('W',f'AH{src}',wval)]:
                ref=f'{col}{r}'
                old_style=cells.get(ref).attrib.get('s') if ref in cells else None
                cc=get_or_create_cell(row, ref, old_style)
                set_formula(cc, f'IF({source_ref}="","",{source_ref})', val if val not in (None,'') else None)
            ref=f'V{r}'
            old_style=cells.get(ref).attrib.get('s') if ref in cells else None
            vc=get_or_create_cell(row, ref, old_style)
            un=numeric(uval); wn=numeric(wval)
            vval = None if tval in (None,'') or un in (None,0) or wval in (None,'') else 1 - (wn or 0)/un
            set_formula(vc, f'IF(OR(U{r}="",U{r}=0,W{r}=""),"",1-W{r}/U{r})', vval)
            updated_hyper+=4
            if tval in (None,''): blank_hyper+=1
        # Add guide block in AP:AQ.
        guide=[
            ('AP1','MAPA - Top Offenders Customers'),
            ('AP2','Fonte'),('AQ2','Todas as pivots desta aba usam ROE_wk como base.'),
            ('AP3','Oficial x Diagnostico'),('AQ3','Week_Overview e o KPI oficial agregam a base por regra OTO; Top Offenders e diagnostico por cliente/filtro.'),
            ('AP4','Quando pode divergir'),('AQ4','Diverge quando filtros de semana, OTO Out, Atrasado, Centro de Custo, cliente, porto ou escopo ficam diferentes.'),
            ('AP5','VISÃO GERAL A:G'),('AQ5','Mostra volume, atrasos e OTO por cliente; coluna G e formula OTO DIA = 1 - atraso/volume.'),
            ('AP6','HYPER CARE NORTE T:W'),('AQ6','Espelho do helper AF:AH; linhas sem helper ficam em branco para nao aparecer 0 falso.'),
            ('AP7','JUSTIFICATIVA Y:AA'),('AQ7','Mostra justificativas de atraso e OTO associado.'),
            ('AP8','Menu'),('AQ8','Menu!C34/E34/G34 puxam AK24; Menu!I34/K34/M34 puxam A6 via GETPIVOTDATA.'),
            ('AP9','Grafico'),('AQ9','Chart do Menu usa AK25:AN43 desta aba.'),
            ('AP10','Regra atual'),('AQ10','Sem Preenchimento conta volume, nao entra no OTO; Especial entra no KPI e nao penaliza quando atrasado.'),
            ('AP11','Checklist'),('AQ11','Antes de comparar com Week_Overview, alinhar semana e filtros oficiais.'),
        ]
        for ref,text in guide:
            r=split_ref(ref)[1]; row=get_or_create_row(sheetData, rows, r); cc=get_or_create_cell(row, ref)
            set_inline(cc,text)
        # Dimension include guide block.
        dim=root.find(q('dimension'))
        if dim is not None: dim.attrib['ref']='A1:AQ510'
        # Add/restore icon sets for OTO columns.
        cf_ranges=['D8:D510','G8:G510','P8:P167','V7:V441','AA7:AA37','AM25:AM43']
        add_icon_cf(root, cf_ranges)
        data=ET.tostring(root, encoding='utf-8', xml_declaration=True)
        tmp=WORKBOOK.with_suffix('.topfix.xlsm')
        with ZipFile(tmp,'w',compression=ZIP_DEFLATED,allowZip64=True) as zout:
            for item in zin.infolist():
                name=item.filename
                if name==sheet_path:
                    zout.writestr(item,data)
                elif name=='xl/calcChain.xml':
                    continue
                else:
                    zout.writestr(item, zin.read(name))
    tmp.replace(WORKBOOK)
    report={
        'workbook':str(WORKBOOK),'backup':str(backup),'sheet':SHEET,
        'g510_formula':'IFERROR(-(C510/B510-1),"")','g510_cached_value':g_value,
        'hypercare_formulas_updated':updated_hyper,'hypercare_blank_source_rows':blank_hyper,
        'conditional_formatting_ranges':cf_ranges,
        'guide_block':'AP1:AQ11',
        'workbook_size':WORKBOOK.stat().st_size,
        'workbook_mtime':datetime.fromtimestamp(WORKBOOK.stat().st_mtime).isoformat(timespec='seconds')
    }
    REPORT.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(report,ensure_ascii=False,indent=2))
    print('REPORT='+str(REPORT))
if __name__=='__main__': main()
