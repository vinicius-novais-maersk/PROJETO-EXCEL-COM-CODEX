
from pathlib import Path
import json, traceback

WB_PATH = Path(r"C:\Users\VNO024\Downloads\Base_DSU2026 - TbM - WK19.xlsm")
OUT = Path(r"C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX\analysis\explanation_maps_update_v2_result.json")
XL_CENTER = -4108
XL_TOP = -4160
XL_CONTINUOUS = 1
XL_THIN = 2

TOP_MAP = [
    ("MAPA - Top Offenders Customers", ""),
    ("Objetivo", "Diagn\u00f3stico: mostrar quais clientes puxam volume, atrasos e OTO. N\u00e3o \u00e9 o KPI oficial; \u00e9 uma aba para explicar o detalhe."),
    ("Base", "Fonte principal: ROE_wk. As tabelas podem ter filtros pr\u00f3prios, ent\u00e3o sempre conferir filtros antes de comparar."),
    ("Compara\u00e7\u00e3o com Week_Overview", "Week_Overview mostra o KPI agregado por porto/regi\u00e3o/BR. Top Offenders quebra por cliente. Para bater, semana, Volume, OTO Out, Atrasado?, Centro de Custo, Porto/Regi\u00e3o e regras OTO precisam estar iguais."),
    ("VIS\u00c3O GERAL A:G", "Volume, atraso e OTO por cliente. D = OTO WK. G = OTO DIA/helper. O Total Geral deve conversar com o volume da semana quando os filtros est\u00e3o iguais."),
    ("ATRASOS N:R", "Recorte dos clientes com atraso. Serve para descobrir quem gerou a queda de OTO."),
    ("HYPER CARE NORTE T:W", "Recorte de Hypercare Norte. Linhas sem dado ficam vazias para evitar zero falso."),
    ("JUSTIFICATIVAS Y:AA", "Agrupa os motivos de atraso. Use para explicar por que o OTO caiu."),
    ("HYPERCARE AK:AN", "Tabela usada pelo gr\u00e1fico do Menu para o recorte Hypercare."),
    ("Menu", "Menu!C34/E34/G34 usa AK24. Menu!I34/K34/M34 usa A6 via GETPIVOTDATA."),
    ("Regras OTO", "Sem Preenchimento conta no volume e sai do KPI. Especial conta no volume/KPI, mas atraso Especial n\u00e3o penaliza."),
    ("Como usar", "Se algu\u00e9m reclamar diferen\u00e7a: primeiro alinhar filtros. Se filtros estiverem iguais, comparar Total BR/porto na Week_Overview."),
]

WEEK_MAP = [
    ("MAPA - Week Overview: colunas principais", "", "", ""),
    ("Coluna", "O que representa", "Como calcula", "Fonte / regra"),
    ("Capacity", "Capacidade planejada para o porto/regi\u00e3o.", "Valor de refer\u00eancia cadastrado por semana/porto; totais s\u00e3o usados como comparativo operacional.", "Base auxiliar Amb+RoFo+Cap."),
    ("OTO CAB (%)", "KPI OTO do CAB / Alian\u00e7a.", "1 - atrasos penaliz\u00e1veis / volume eleg\u00edvel CAB.", "ROE_wk: Volume=Ok, Centro de Custo=Alian\u00e7a, OTO Out=N. Sem Preenchimento fica fora do KPI. Especial atrasado n\u00e3o penaliza."),
    ("CTO (Customer Time Operation)", "Vis\u00e3o de impacto relacionado a tempo/cliente para CAB.", "Usa o OTO CAB ponderado pelos moves CAB e causas cliente: N * D / (D + SUM(W:AA)).", "Pode ser diferente do OTO CAB porque considera causas/ajustes de cliente."),
    ("OTO DS (%)", "KPI OTO do DS / n\u00e3o-Alian\u00e7a.", "Mesma regra do OTO, mas com Centro de Custo <> Alian\u00e7a.", "Sem Preenchimento fora do KPI; Especial atrasado n\u00e3o penaliza."),
    ("OTO Ttl (%)", "KPI OTO total da linha/porto/regi\u00e3o/BR.", "1 - atrasos penaliz\u00e1veis / volume eleg\u00edvel total.", "Todos os centros de custo; Sem Preenchimento fora do KPI; Especial atrasado n\u00e3o penaliza."),
    ("48h Schedule CAB (%)", "% de CAB agendado dentro da regra de 48h.", "1 - quantidade com SLA Ag=N / Moves CAB.", "ROE_wk: Volume=Ok, Centro de Custo=Alian\u00e7a, SLA Ag."),
    ("Quando n\u00e3o bater", "Normalmente \u00e9 filtro diferente, n\u00e3o erro de conta.", "Conferir Semana, Porto/Regi\u00e3o, Volume, OTO Out, Atrasado?, Centro de Custo e regra Especial/Sem Preenchimento.", "Piv\u00f4 com filtro travado distorce a compara\u00e7\u00e3o."),
]

REGION_NOTE = [
    ("MAPA - Region errors", ""),
    ("\u00c9 tabela din\u00e2mica?", "N\u00e3o. Esta aba \u00e9 uma tabela auxiliar por f\u00f3rmula, n\u00e3o uma tabela din\u00e2mica."),
    ("Dias da semana", "B6:H6 s\u00e3o crit\u00e9rios usados nas f\u00f3rmulas COUNTIFS contra ROE_wk[day week]."),
    ("Atualiza\u00e7\u00e3o", "Quando ROE_wk e Week_Overview!AG1 mudam/recalculam, os dias atualizam sem depender de cache/filtro de piv\u00f4."),
]

def apply_table(ws, start_cell, rows, title_fill=0x1F4E78, header_fill=0xD9EAF7):
    start = ws.Range(start_cell)
    r0 = start.Row
    c0 = start.Column
    nrows = len(rows)
    ncols = max(len(r) for r in rows)
    rng = ws.Range(ws.Cells(r0, c0), ws.Cells(r0+nrows-1, c0+ncols-1))
    try: rng.UnMerge()
    except Exception: pass
    rng.ClearContents()
    rng.WrapText = True
    rng.VerticalAlignment = XL_TOP
    for i,row in enumerate(rows):
        for j,val in enumerate(row):
            ws.Cells(r0+i, c0+j).Value = val
    title_rng = ws.Range(ws.Cells(r0,c0), ws.Cells(r0,c0+ncols-1))
    try: title_rng.Merge()
    except Exception: pass
    title_rng.Font.Bold = True
    title_rng.Font.Color = 0xFFFFFF
    title_rng.Interior.Color = title_fill
    title_rng.HorizontalAlignment = XL_CENTER
    if ncols > 2 and nrows > 1:
        header_rng = ws.Range(ws.Cells(r0+1,c0), ws.Cells(r0+1,c0+ncols-1))
        header_rng.Font.Bold = True
        header_rng.Interior.Color = header_fill
    first_col = ws.Range(ws.Cells(r0,c0), ws.Cells(r0+nrows-1,c0))
    first_col.Font.Bold = True
    for b in range(7, 13):
        try:
            rng.Borders(b).LineStyle = XL_CONTINUOUS
            rng.Borders(b).Weight = XL_THIN
            rng.Borders(b).Color = 0xB7B7B7
        except Exception:
            pass
    return rng.Address

def main():
    import win32com.client as win32
    xl = None; wb = None
    result = {"workbook": str(WB_PATH), "updates": [], "errors": []}
    try:
        xl = win32.DispatchEx('Excel.Application')
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.EnableEvents = False
        try: xl.ScreenUpdating = False
        except Exception: pass
        wb = xl.Workbooks.Open(str(WB_PATH), UpdateLinks=0, ReadOnly=False)
        ws_top = wb.Worksheets('Top_Offenders_Customers')
        result['updates'].append({'sheet':'Top_Offenders_Customers','range':apply_table(ws_top,'AP1',TOP_MAP)})
        ws_top.Columns('AP').ColumnWidth = 30
        ws_top.Columns('AQ').ColumnWidth = 105
        ws_top.Rows('1:12').AutoFit()
        ws_week = wb.Worksheets('Week_Overview')
        result['updates'].append({'sheet':'Week_Overview','range':apply_table(ws_week,'S25',WEEK_MAP)})
        ws_week.Columns('S').ColumnWidth = 30
        ws_week.Columns('T').ColumnWidth = 44
        ws_week.Columns('U').ColumnWidth = 62
        ws_week.Columns('V').ColumnWidth = 72
        ws_week.Rows('25:33').AutoFit()
        ws_region = wb.Worksheets('Region errors')
        result['updates'].append({'sheet':'Region errors','range':apply_table(ws_region,'A14',REGION_NOTE,title_fill=0x70AD47,header_fill=0xE2F0D9)})
        ws_region.Columns('A').ColumnWidth = 26
        ws_region.Columns('B').ColumnWidth = 100
        ws_region.Rows('14:17').AutoFit()
        wb.Save()
        result['saved'] = True
    except Exception as e:
        result['errors'].append({'error': repr(e), 'traceback': traceback.format_exc()})
        raise
    finally:
        try:
            if wb is not None: wb.Close(SaveChanges=False)
        except Exception: pass
        try:
            if xl is not None: xl.Quit()
        except Exception: pass
        OUT.write_text(json.dumps(result, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
        print(json.dumps(result, ensure_ascii=False, indent=2, default=str))

if __name__ == '__main__':
    main()
