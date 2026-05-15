
from pathlib import Path
import json, time, traceback

WB_PATH = Path(r"C:\Users\VNO024\Downloads\Base_DSU2026 - TbM - WK19.xlsm")
OUT = Path(r"C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX\analysis\final_audit_fix_result.json")
ALI = 'Alian' + chr(0xE7) + 'a'

formula_map = {
    ('Volume_DS','M11'): f'=COUNTIFS(\'ROE_wk\'!$AV:$AV,"Ok",\'ROE_wk\'!$AY:$AY,$M$3,\'ROE_wk\'!$AR:$AR,$M$2,\'ROE_wk\'!$AI:$AI,"{ALI}")',
    ('Volume_DS','M12'): f'=COUNTIFS(\'ROE_wk\'!$AV:$AV,"Ok",\'ROE_wk\'!$AY:$AY,$M$3,\'ROE_wk\'!$AR:$AR,$M$2,\'ROE_wk\'!$AI:$AI,"<>{ALI}")',
    ('Volume_MAO','M12'): f'=COUNTIFS(\'ROE_wk\'!$AV:$AV,"Ok",\'ROE_wk\'!$AY:$AY,$M$3,\'ROE_wk\'!$AR:$AR,$M$2,\'ROE_wk\'!$AI:$AI,"{ALI}",\'ROE_wk\'!$BO:$BO,"<>Especial")',
    ('Volume_MAO','M13'): f'=COUNTIFS(\'ROE_wk\'!$AV:$AV,"Ok",\'ROE_wk\'!$AY:$AY,$M$3,\'ROE_wk\'!$AR:$AR,$M$2,\'ROE_wk\'!$AI:$AI,"{ALI}",\'ROE_wk\'!$BO:$BO,"Especial")',
    ('Volume_MAO','M14'): f'=COUNTIFS(\'ROE_wk\'!$AV:$AV,"Ok",\'ROE_wk\'!$AY:$AY,$M$3,\'ROE_wk\'!$AR:$AR,$M$2,\'ROE_wk\'!$AI:$AI,"<>{ALI}")',
    ('Volume_Graph','M12'): f'=COUNTIFS(\'ROE_wk\'!$AV:$AV,"Ok",\'ROE_wk\'!$AY:$AY,$M$3,\'ROE_wk\'!$AR:$AR,$M$2,\'ROE_wk\'!$AI:$AI,"{ALI}")',
    ('Volume_Graph','M13'): f'=COUNTIFS(\'ROE_wk\'!$AV:$AV,"Ok",\'ROE_wk\'!$AY:$AY,$M$3,\'ROE_wk\'!$AR:$AR,$M$2,\'ROE_wk\'!$AI:$AI,"<>{ALI}")',
}

def field_by_name(pt, name):
    # Exact first, then case-insensitive trimmed fallback.
    try:
        return pt.PivotFields(name)
    except Exception:
        wanted = name.strip().lower()
        for i in range(1, pt.PivotFields().Count + 1):
            pf = pt.PivotFields(i)
            if str(pf.Name).strip().lower() == wanted:
                return pf
    raise KeyError(name)

def clear_filter(pt, name):
    pf = field_by_name(pt, name)
    try:
        pf.ClearAllFilters()
    except Exception:
        pass
    return pf

def set_page(pt, name, value=None):
    pf = clear_filter(pt, name)
    if value is not None:
        last_err = None
        for val in (value, str(value)):
            try:
                pf.CurrentPage = val
                return True
            except Exception as e:
                last_err = e
        raise last_err
    return True

def refresh_pivot(pt):
    try:
        pt.PivotCache().Refresh()
    except Exception:
        pass
    try:
        pt.RefreshTable()
    except Exception:
        pass

def main():
    import win32com.client as win32
    xl = None
    wb = None
    result = {'workbook': str(WB_PATH), 'formula_updates': [], 'pivot_updates': [], 'validation': {}, 'errors': []}
    try:
        xl = win32.DispatchEx('Excel.Application')
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.EnableEvents = False
        try:
            xl.AskToUpdateLinks = False
        except Exception:
            pass
        wb = xl.Workbooks.Open(str(WB_PATH), UpdateLinks=0, ReadOnly=False)
        # Fix formula mojibake in total cells.
        for (sheet, cell), formula in formula_map.items():
            ws = wb.Worksheets(sheet)
            before_formula = ws.Range(cell).Formula
            ws.Range(cell).Formula = formula
            result['formula_updates'].append({'sheet': sheet, 'cell': cell, 'before': str(before_formula), 'after': formula})
        # Refresh/set Top_Offenders_Vendors filters based on old known-good state.
        ws = wb.Worksheets('Top_Offenders_Vendors')
        pts = ws.PivotTables()
        for i in range(1, pts.Count + 1):
            pt = pts.Item(i)
            try:
                top_col = int(pt.TableRange2.Column)
                name = str(pt.Name)
                addr = str(pt.TableRange2.Address)
                # Refresh cache/table before applying filters.
                refresh_pivot(pt)
                if top_col <= 5:
                    # VIS?O GERAL: all OTO/Atrasado, Volume Ok, all Data Prog.
                    set_page(pt, 'Volume', 'Ok')
                    set_page(pt, 'Data Prog.', None)
                    set_page(pt, 'OTO Out', None)
                    set_page(pt, 'Atrasado?', None)
                    role = 'VIS?O GERAL'
                elif 11 <= top_col <= 14:
                    # ATRASOS: only delayed, all OTO Out.
                    set_page(pt, 'Volume', 'Ok')
                    set_page(pt, 'Data Prog.', None)
                    set_page(pt, 'Atrasado?', '1')
                    set_page(pt, 'OTO Out', None)
                    role = 'ATRASOS'
                else:
                    # JUSTIFICATIVAS: only delayed, all OTO Out.
                    set_page(pt, 'Volume', 'Ok')
                    set_page(pt, 'Data Prog.', None)
                    set_page(pt, 'Atrasado?', '1')
                    set_page(pt, 'OTO Out', None)
                    role = 'JUSTIFICATIVAS'
                refresh_pivot(pt)
                result['pivot_updates'].append({'name': name, 'role': role, 'initial_address': addr, 'top_col': top_col, 'status': 'updated'})
            except Exception as e:
                result['pivot_updates'].append({'name': getattr(pt, 'Name', '?'), 'status': 'error', 'error': repr(e)})
        # Full calculation and chart refresh.
        try:
            xl.CalculateFullRebuild()
        except Exception:
            xl.CalculateFull()
        # Validation snapshot.
        for sheet, cell in formula_map:
            ws2 = wb.Worksheets(sheet)
            result['validation'][f'{sheet}!{cell}'] = {'formula': str(ws2.Range(cell).Formula), 'value': ws2.Range(cell).Value}
        wv = wb.Worksheets('Top_Offenders_Vendors')
        result['validation']['Top_Offenders_Vendors_filters'] = {
            'B2': wv.Range('B2').Value, 'B3': wv.Range('B3').Value, 'B4': wv.Range('B4').Value, 'B5': wv.Range('B5').Value,
            'M2': wv.Range('M2').Value, 'M3': wv.Range('M3').Value, 'M4': wv.Range('M4').Value, 'M5': wv.Range('M5').Value,
            'P2': wv.Range('P2').Value, 'P3': wv.Range('P3').Value, 'P4': wv.Range('P4').Value, 'P5': wv.Range('P5').Value,
        }
        result['validation']['Top_Offenders_Vendors_rows'] = {
            'A8:E15': [[wv.Cells(r,c).Value for c in range(1,6)] for r in range(8,16)],
            'L8:M15': [[wv.Cells(r,c).Value for c in range(12,14)] for r in range(8,16)],
            'O8:P15': [[wv.Cells(r,c).Value for c in range(15,17)] for r in range(8,16)],
        }
        wb.Save()
        result['saved'] = True
    except Exception as e:
        result['errors'].append({'error': repr(e), 'traceback': traceback.format_exc()})
        raise
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if xl is not None:
                xl.Quit()
        except Exception:
            pass
        OUT.write_text(json.dumps(result, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
        print(json.dumps(result, ensure_ascii=False, indent=2, default=str))

if __name__ == '__main__':
    main()
