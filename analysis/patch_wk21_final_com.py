import json
from collections import Counter, defaultdict
from pathlib import Path
import zipfile
import win32com.client as win32

path = Path(r'C:\Users\VNO024\OneDrive - Maersk Group\2025\Py\DSU 2025\Base_DSU2026 - TbM - WK21.xlsm')
report_path = Path(r'c:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX\analysis\onedrive_wk21_final_com_patch_report.json')

# Excel constants
XL_UP = -4162
XL_TO_LEFT = -4159
XL_CALC_AUTO = -4105

# ASCII-safe text helpers. This avoids mojibake in PowerShell/Python source encoding.
def u(code):
    return chr(code)

VIS_GERAL = 'VIS' + u(0x00C3) + 'O GERAL'
ROTULOS = 'R' + u(0x00F3) + 'tulos de Linha'
E_TABELA = u(0x00C9) + ' tabela din' + u(0x00E2) + 'mica?'
NAO_MAPA = 'N' + u(0x00E3) + 'o. Esta aba agora usa uma tabela auxiliar calculada contra ROE_wk.'
CRITERIOS = 'Crit' + u(0x00E9) + 'rios'
ATUALIZACAO = 'Atualiza' + u(0x00E7) + u(0x00E3) + 'o'
PIVO = 'piv' + u(0x00F4)
FORMULAS = 'f' + u(0x00F3) + 'rmulas'
DSU_NOTE_REGION = 'Weeknum=21; Volume=Ok; OTO Out=N; OTD ajustado=Sem Preenchimento.'
DSU_NOTE_REFRESH = 'Valores recalculados no arquivo WK21 para aparecer corretamente sem depender do cache de ' + PIVO + '.'
WEEKLY_LABEL = 'Vol x RoFo (Weekly) atualizado'
QTD_SEM = 'Qtd OS sem preenchimento'
TOTAL_GERAL = 'Total Geral'

SEM_PREENCH = 'Sem Preenchimento'
ATRASADO = 'Atrasado'
ESPECIAL = 'Especial'
OK = 'Ok'
NO = 'N'


def norm(v):
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def num(v):
    if v is None or v == '':
        return 0.0
    try:
        return float(v)
    except Exception:
        return 0.0


def week_num(v):
    if v is None or v == '':
        return None
    try:
        return int(float(v))
    except Exception:
        return None


def cell(ws, row, col, value):
    ws.Cells(row, col).Value = value


def clear_pivots(ws):
    errors = []
    try:
        count = ws.PivotTables().Count
    except Exception:
        return errors
    for i in range(count, 0, -1):
        try:
            pt = ws.PivotTables().Item(i)
            addr = pt.TableRange2.Address
            pt.TableRange2.Clear()
            errors.append({'pivot': str(pt.Name), 'cleared': addr})
        except Exception as exc:
            errors.append({'pivot_index': i, 'error': str(exc)})
    return errors


def write_block(ws, start_row, start_col, rows):
    for r_offset, row in enumerate(rows):
        for c_offset, value in enumerate(row):
            cell(ws, start_row + r_offset, start_col + c_offset, value)


def apply_basic_format(ws_top, ws_region):
    # Top offenders formatting
    for rng in ['A1:E1', 'L1:M1', 'O1:P1', 'R1:S1']:
        try:
            ws_top.Range(rng).Font.Bold = True
            ws_top.Range(rng).Interior.Color = 15773696
        except Exception:
            pass
    for rng in ['A7:E7', 'L7:M7', 'O7:P7', 'R7:S7']:
        try:
            ws_top.Range(rng).Font.Bold = True
            ws_top.Range(rng).Interior.Color = 14277081
        except Exception:
            pass
    try:
        ws_top.Range('D:E').NumberFormat = '0.0%'
        ws_top.Range('A:S').Columns.AutoFit()
    except Exception:
        pass

    # Region errors formatting
    for rng in ['A1:B4', 'A6:J6', 'A14:B17']:
        try:
            ws_region.Range(rng).Font.Bold = True
        except Exception:
            pass
    try:
        ws_region.Range('A6:J6').Interior.Color = 14277081
        ws_region.Range('A:J').Columns.AutoFit()
    except Exception:
        pass


def refresh_charts(wb):
    refreshed = []
    for ws_name in ['Menu', 'Volume_Graph', 'Volume_DS', 'Volume_MAO']:
        try:
            ws = wb.Worksheets(ws_name)
            count = ws.ChartObjects().Count
        except Exception:
            continue
        for i in range(1, count + 1):
            try:
                co = ws.ChartObjects(i)
                ch = co.Chart
                sc = ch.SeriesCollection().Count
                for j in range(1, sc + 1):
                    s = ch.SeriesCollection(j)
                    try:
                        formula = s.Formula
                        s.Formula = formula
                    except Exception:
                        pass
                try:
                    ch.Refresh()
                except Exception:
                    pass
                refreshed.append(ws_name + '!' + str(co.Name))
            except Exception as exc:
                refreshed.append(ws_name + '!chart_' + str(i) + ' ERROR ' + str(exc))
    return refreshed


def check_zip_ok(path):
    with zipfile.ZipFile(path, 'r') as zf:
        bad = zf.testzip()
        has_vba = 'xl/vbaProject.bin' in zf.namelist()
    return {'zip_bad_file': bad, 'has_vbaProject_bin': has_vba}

excel = win32.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
excel.AskToUpdateLinks = False
excel.EnableEvents = False
try:
    excel.Calculation = XL_CALC_AUTO
except Exception:
    pass

wb = None
report = {'workbook': str(path)}
try:
    wb = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True, AddToMru=False)
    try:
        target_week = week_num(wb.Worksheets('Week_Overview').Range('AG1').Value) or 21
    except Exception:
        target_week = 21
    report['target_week'] = target_week

    roe = wb.Worksheets('ROE_wk')
    values = roe.UsedRange.Value
    headers = [norm(x) for x in values[0]]
    idx = {h: i for i, h in enumerate(headers)}
    required = ['Weeknum', 'Volume', 'Provedor', 'AtrasoRev', '% OTO Provider', 'OTD ajustado', 'OTO Out', 'Especiais', 'Justificativa', 'Region', 'day week']
    missing = [h for h in required if h not in idx]
    if missing:
        raise RuntimeError('Missing required headers: ' + ', '.join(missing))

    data_rows = values[1:]
    wk_rows = [r for r in data_rows if week_num(r[idx['Weeknum']]) == target_week]
    volume_ok_rows = [r for r in wk_rows if norm(r[idx['Volume']]).lower() == OK.lower()]

    # Vendor overview table
    prov_stats = {}
    for r in volume_ok_rows:
        provider = norm(r[idx['Provedor']]) or '(blank)'
        if provider not in prov_stats:
            prov_stats[provider] = {'count': 0, 'delay': 0.0, 'oto_sum': 0.0, 'oto_n': 0}
        d = prov_stats[provider]
        d['count'] += 1
        d['delay'] += num(r[idx['AtrasoRev']])
        oto = r[idx['% OTO Provider']]
        if oto not in (None, ''):
            d['oto_sum'] += num(oto)
            d['oto_n'] += 1

    provider_rows = []
    for provider, d in sorted(prov_stats.items(), key=lambda kv: (-kv[1]['count'], kv[0])):
        count = d['count']
        delay = d['delay']
        oto_wk = (d['oto_sum'] / d['oto_n']) if d['oto_n'] else None
        oto_dia = (1 - delay / count) if count else None
        provider_rows.append([provider, count, delay, oto_wk, oto_dia])
    total_count = sum(d['count'] for d in prov_stats.values())
    total_delay = sum(d['delay'] for d in prov_stats.values())
    total_oto_sum = sum(d['oto_sum'] for d in prov_stats.values())
    total_oto_n = sum(d['oto_n'] for d in prov_stats.values())
    provider_total = [TOTAL_GERAL, total_count, total_delay, (total_oto_sum / total_oto_n) if total_oto_n else None, (1 - total_delay / total_count) if total_count else None]

    delayed_rows = [
        r for r in volume_ok_rows
        if norm(r[idx['OTD ajustado']]) == ATRASADO and norm(r[idx['OTO Out']]).upper() == NO and norm(r[idx['Especiais']]) != ESPECIAL
    ]
    delay_by_provider = Counter(norm(r[idx['Provedor']]) or '(blank)' for r in delayed_rows)
    delay_provider_rows = [[k, v] for k, v in sorted(delay_by_provider.items(), key=lambda kv: (-kv[1], kv[0]))]
    delay_provider_total = [TOTAL_GERAL, sum(delay_by_provider.values())]

    just_counts = Counter(norm(r[idx['Justificativa']]) or '(blank)' for r in delayed_rows)
    just_rows = [[k, v] for k, v in sorted(just_counts.items(), key=lambda kv: (-kv[1], kv[0]))]
    just_total = [TOTAL_GERAL, sum(just_counts.values())]

    # Region errors table
    region_source_rows = [
        r for r in volume_ok_rows
        if norm(r[idx['OTD ajustado']]) == SEM_PREENCH and norm(r[idx['OTO Out']]).upper() == NO
    ]
    day_order = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']
    region_stats = defaultdict(Counter)
    for r in region_source_rows:
        region = norm(r[idx['Region']]) or 'Sem Porto/Region'
        day = norm(r[idx['day week']]) or '(blank)'
        region_stats[region][day] += 1
    region_rows = []
    for region, counter in sorted(region_stats.items(), key=lambda kv: (-sum(kv[1].values()), kv[0])):
        vals = [counter.get(day, 0) for day in day_order]
        region_rows.append([region, sum(vals)] + vals + [sum(vals)])
    total_day_vals = [sum(counter.get(day, 0) for counter in region_stats.values()) for day in day_order]
    region_total = [TOTAL_GERAL, sum(total_day_vals)] + total_day_vals + [sum(total_day_vals)]

    # Patch Top_Offenders_Vendors
    ws_top = wb.Worksheets('Top_Offenders_Vendors')
    top_pivots = clear_pivots(ws_top)
    ws_top.Range('A1:S200').Clear()
    cell(ws_top, 1, 1, VIS_GERAL)
    cell(ws_top, 1, 12, 'ATRASOS')
    cell(ws_top, 1, 15, 'JUSTIFICATIVAS')
    cell(ws_top, 1, 18, 'RESUMO JUSTIFICATIVAS')
    # Filter captions
    for start_col in [1, 12, 15, 18]:
        cell(ws_top, 2, start_col, 'Volume')
        cell(ws_top, 2, start_col + 1, OK)
        cell(ws_top, 3, start_col, 'Data Prog.')
        cell(ws_top, 3, start_col + 1, '(Tudo)')
    cell(ws_top, 4, 1, 'OTO Out')
    cell(ws_top, 4, 2, '(Tudo)')
    cell(ws_top, 5, 1, 'Atrasado?')
    cell(ws_top, 5, 2, '(Tudo)')
    for start_col in [12, 15, 18]:
        cell(ws_top, 4, start_col, 'Atrasado?')
        cell(ws_top, 4, start_col + 1, 1)
        cell(ws_top, 5, start_col, 'OTO Out')
        cell(ws_top, 5, start_col + 1, NO)

    write_block(ws_top, 7, 1, [[ROTULOS, 'Count of OS', 'Sum of AtrasoRev', 'OTO WK', 'OTO DIA']])
    write_block(ws_top, 8, 1, provider_rows + [provider_total])
    write_block(ws_top, 7, 12, [[ROTULOS, 'Count of OS']])
    write_block(ws_top, 8, 12, delay_provider_rows + [delay_provider_total])
    write_block(ws_top, 7, 15, [[ROTULOS, 'Count of OS']])
    write_block(ws_top, 8, 15, just_rows + [just_total])
    write_block(ws_top, 7, 18, [[ROTULOS, 'Count of OS']])
    write_block(ws_top, 8, 18, just_rows + [just_total])

    # Patch Region errors
    ws_region = wb.Worksheets('Region errors')
    region_pivots = clear_pivots(ws_region)
    ws_region.Range('A1:J40').Clear()
    write_block(ws_region, 1, 1, [
        ['OTO Out', NO],
        ['OTD ajustado', SEM_PREENCH],
        ['Volume', OK],
        ['Weeknum', target_week],
    ])
    write_block(ws_region, 6, 1, [['Region', QTD_SEM] + day_order + ['Total dias']])
    write_block(ws_region, 7, 1, region_rows + [region_total])
    write_block(ws_region, 14, 1, [
        ['MAPA - Region errors', None],
        [E_TABELA, NAO_MAPA],
        [CRITERIOS, DSU_NOTE_REGION],
        [ATUALIZACAO, DSU_NOTE_REFRESH],
    ])

    # Menu/source chart cleanup labels
    for sheet_name, addr in [('Volume_Graph', 'E28'), ('Volume_DS', 'E28'), ('Volume_MAO', 'E30')]:
        try:
            wb.Worksheets(sheet_name).Range(addr).Value = WEEKLY_LABEL
        except Exception as exc:
            report.setdefault('label_errors', []).append(sheet_name + '!' + addr + ': ' + str(exc))

    apply_basic_format(ws_top, ws_region)

    try:
        excel.CalculateFullRebuild()
    except Exception as exc:
        report['calculate_full_rebuild_error'] = str(exc)
    report['refreshed_charts'] = refresh_charts(wb)

    wb.Save()
    report.update({
        'top_pivots_cleared': top_pivots,
        'region_pivots_cleared': region_pivots,
        'wk_rows': len(wk_rows),
        'volume_ok_rows': len(volume_ok_rows),
        'provider_rows': len(provider_rows),
        'provider_total_count': total_count,
        'provider_total_delay': total_delay,
        'delay_rows': len(delayed_rows),
        'justification_total': sum(just_counts.values()),
        'region_error_rows': len(region_source_rows),
        'region_total': region_total,
        'top_first_row': provider_rows[0] if provider_rows else None,
    })
finally:
    if wb is not None:
        wb.Close(False)
    try:
        excel.EnableEvents = True
    except Exception:
        pass
    excel.Quit()

report['zip'] = check_zip_ok(path)
report_path.parent.mkdir(parents=True, exist_ok=True)
report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps(report, ensure_ascii=True, indent=2))
