
from pathlib import Path
import json, traceback

WB_PATH = Path(r"C:\Users\VNO024\Downloads\Base_DSU2026 - TbM - WK19.xlsm")
OUT = Path(r"C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX\analysis\explanation_maps_update_result.json")

# Excel constants
XL_CENTER = -4108
XL_TOP = -4160
XL_CONTINUOUS = 1
XL_THIN = 2

TOP_MAP = [
    ("MAPA - Top Offenders Customers", ""),
    ("Objetivo", "Diagnostico: mostrar quais clientes puxam volume, atrasos e OTO. Nao e o KPI oficial; e uma aba para explicar o detalhe."),
    ("Base", "Fonte principal: ROE_wk. As tabelas podem ter filtros proprios, entao sempre conferir filtros antes de comparar."),
    ("Comparacao com Week_Overview", "Week_Overview mostra o KPI agregado por porto/regiao/BR. Top Offenders quebra por cliente. Para bater, semana, Volume, OTO Out, Atrasado?, Centro de Custo, Porto/Regiao e regras OTO precisam estar iguais."),
    ("VISAO GERAL A:G", "Volume, atraso e OTO por cliente. D = OTO WK. G = OTO DIA/helper. O Total Geral deve conversar com o volume da semana quando os filtros estao iguais."),
    ("ATRASOS N:R", "Recorte dos clientes com atraso. Serve para descobrir quem gerou a queda de OTO."),
    ("HYPER CARE NORTE T:W", "Recorte de Hypercare Norte. Linhas sem dado ficam vazias para evitar zero falso."),
    ("JUSTIFICATIVAS Y:AA", "Agrupa os motivos de atraso. Use para explicar por que o OTO caiu."),
    ("HYPERCARE AK:AN", "Tabela usada pelo grafico do Menu para o recorte Hypercare."),
    ("Menu", "Menu!C34/E34/G34 usa AK24. Menu!I34/K34/M34 usa A6 via GETPIVOTDATA."),
    ("Regras OTO", "Sem Preenchimento conta no volume e sai do KPI. Especial conta no volume/KPI, mas atraso Especial nao penaliza."),
    ("Como usar", "Se alguem reclamar diferenca: primeiro alinhar filtros. Se filtros estiverem iguais, comparar Total BR/porto na Week_Overview."),
]

WEEK_MAP = [
    ("MAPA - Week Overview: colunas principais", "", "", ""),
    ("Coluna", "O que representa", "Como calcula", "Fonte / regra"),
    ("Capacity", "Capacidade planejada para o porto/regiao.", "Valor de referencia cadastrado por semana/porto; totais sao usados como comparativo operacional.", "Base auxiliar Amb+RoFo+Cap."),
    ("OTO CAB (%)", "KPI OTO do CAB / Alianca.", "1 - atrasos penalizaveis / volume elegivel CAB.", "ROE_wk: Volume=Ok, Centro de Custo=Alianca, OTO Out=N. Sem Preenchimento fica fora do KPI. Especial atrasado nao penaliza."),
    ("CTO (Customer Time Operation)", "Visao de impacto relacionado a tempo/cliente para CAB.", "Usa o OTO CAB ponderado pelos moves CAB e causas cliente: N * D / (D + SUM(W:AA)).", "Pode ser diferente do OTO CAB porque considera causas/ajustes de cliente."),
    ("OTO DS (%)", "KPI OTO do DS / nao-Alianca.", "Mesma regra do OTO, mas com Centro de Custo <> Alianca.", "Sem Preenchimento fora do KPI; Especial atrasado nao penaliza."),
    ("OTO Ttl (%)", "KPI OTO total da linha/porto/regiao/BR.", "1 - atrasos penalizaveis / volume elegivel total.", "Todos os centros de custo; Sem Preenchimento fora do KPI; Especial atrasado nao penaliza."),
    ("48h Schedule CAB (%)", "% de CAB agendado dentro da regra de 48h.", "1 - quantidade com SLA Ag=N / Moves CAB.", "ROE_wk: Volume=Ok, Centro de Custo=Alianca, SLA Ag."),
    ("Quando nao bater", "Normalmente e filtro diferente, nao erro de conta.", "Conferir Semana, Porto/Regiao, Volume, OTO Out, Atrasado?, Centro de Custo e regra Especial/Sem Preenchimento.", "Pivo com filtro travado distorce a comparacao."),
]

REGION_NOTE = [
    ("Observacao", "Esta aba nao e uma tabela dinamica; e uma tabela auxiliar por formula."),
    ("Dias da semana", "B6:H6 sao criterios usados nas formulas COUNTIFS contra ROE_wk[day week]."),
    ("Atualizacao", "Quando ROE_wk e Week_Overview!AG1 mudam/recalculam, os dias atualizam sem depender de cache/filtro de pivo."),
]

def apply_table(ws, start_cell, rows, title_fill=0x1F4E78, header_fill=0xD9EAF7):
    start = ws.Range(start_cell)
    r0 = start.Row
    c0 = start.Column
    nrows = len(rows)
    ncols = max(len(r) for r in rows)
    rng = ws.Range(ws.Cells(r0, c0), ws.Cells(r0+nrows-1, c0+ncols-1))
    try:
        rng.UnMerge()
    except Exception:
        pass
    rng.ClearContents()
    rng.WrapText = True
    rng.VerticalAlignment = XL_TOP
    for i,row in enumerate(rows):
        for j,val in enumerate(row):
            ws.Cells(r0+i, c0+j).Value = val
    # title row
    title_rng = ws.Range(ws.Cells(r0,c0), ws.Cells(r0,c0+ncols-1))
    try:
        title_rng.Merge()
    except Exception:
        pass
    title_rng.Font.Bold = True
    title_rng.Font.Color = 0xFFFFFF
    title_rng.Interior.Color = title_fill
    title_rng.HorizontalAlignment = XL_CENTER
    # second row as header for 4-col map only
    if ncols > 2 and nrows > 1:
        header_rng = ws.Range(ws.Cells(r0+1,c0), ws.Cells(r0+1,c0+ncols-1))
        header_rng.Font.Bold = True
        header_rng.Interior.Color = header_fill
    # first column bold
    first_col = ws.Range(ws.Cells(r0,c0), ws.Cells(r0+nrows-1,c0))
    first_col.Font.Bold = True
    # borders
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
    xl = None
    wb = None
    result = {"workbook": str(WB_PATH), "updates": [], "errors": []}
    try:
        xl = win32.DispatchEx('Excel.Application')
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.EnableEvents = False
        wb = xl.Workbooks.Open(str(WB_PATH), UpdateLinks=0, ReadOnly=False)

        ws_top = wb.Worksheets('Top_Offenders_Customers')
        addr = apply_table(ws_top, 'AP1', TOP_MAP)
        ws_top.Columns('AP').ColumnWidth = 28
        ws_top.Columns('AQ').ColumnWidth = 100
        ws_top.Rows('1:12').AutoFit()
        result['updates'].append({'sheet': 'Top_Offenders_Customers', 'range': addr, 'description': 'MAPA reescrito com explicacao mais clara'})

        ws_week = wb.Worksheets('Week_Overview')
        addr = apply_table(ws_week, 'S25', WEEK_MAP)
        ws_week.Columns('S').ColumnWidth = 28
        ws_week.Columns('T').ColumnWidth = 42
        ws_week.Columns('U').ColumnWidth = 58
        ws_week.Columns('V').ColumnWidth = 68
        ws_week.Rows('25:33').AutoFit()
        result['updates'].append({'sheet': 'Week_Overview', 'range': addr, 'description': 'MAPA criado para Capacity, OTO CAB, CTO, OTO DS, OTO Ttl e 48h Schedule CAB'})

        ws_region = wb.Worksheets('Region errors')
        addr = apply_table(ws_region, 'A14', REGION_NOTE, title_fill=0x70AD47, header_fill=0xE2F0D9)
        ws_region.Columns('A').ColumnWidth = 24
        ws_region.Columns('B').ColumnWidth = 95
        ws_region.Rows('14:16').AutoFit()
        result['updates'].append({'sheet': 'Region errors', 'range': addr, 'description': 'Nota adicionada explicando que dias sao formulas, nao pivo'})

        try:
            xl.CalculateFullRebuild()
        except Exception:
            xl.CalculateFull()
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
