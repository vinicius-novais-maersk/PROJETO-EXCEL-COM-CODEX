from __future__ import annotations

from pathlib import Path

import pythoncom
import win32com.client

try:
    from scripts.workbook_paths import resolve_workbook_path
except ModuleNotFoundError:
    from workbook_paths import resolve_workbook_path

WORKBOOK_PATH = resolve_workbook_path()
FALLBACK_SHEET = "PortRegion_Fallback"
ROE_SHEET = "ROE_wk"
TABLE_NAME = "ROE_wk"


def main() -> None:
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False

    workbook = None
    try:
        workbook = excel.Workbooks.Open(str(WORKBOOK_PATH), UpdateLinks=0, ReadOnly=False)

        ws = workbook.Worksheets(FALLBACK_SHEET)
        roe = workbook.Worksheets(ROE_SHEET)
        table = roe.ListObjects(TABLE_NAME)
        last_fallback_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
        last_roe_row = roe.Cells(roe.Rows.Count, 1).End(-4162).Row

        # Move the static fallback map to hidden columns so the visible section can become dynamic.
        ws.Range(f"L1:T{last_fallback_row}").Value = ws.Range(f"A1:I{last_fallback_row}").Value

        # Clear visible area and rebuild it as a dynamic monitoring view.
        ws.Range("A1:J250").Clear()

        ws.Range("A1:I1").Merge()
        ws.Range("A1").Value = "Fallback de Porto/Região - Casos Ativos"

        ws.Range("A2:I2").Merge()
        ws.Range("A2").Value = "Esta aba mostra apenas OS que ainda estão sem classificação oficial no RAO_wk."

        ws.Range("A3:I3").Merge()
        ws.Range("A3").Value = "Se a OS aparecer corretamente no RAO_wk, ela sai desta lista automaticamente."

        ws.Range("A4:I4").Merge()
        ws.Range("A4").Value = "A base técnica usada para inferência continua nesta aba, em colunas ocultas de apoio."

        ws.Range("A6").Value = "Casos ativos"
        ws.Range("C6").Value = "Sem Match"
        ws.Range("E6").Value = "Inferidos"

        ws.Range("A7").Formula = '=COUNTIFS(ROE_wk!$A$2:$A$9915,"<>",ROE_wk!$AV$2:$AV$9915,"Ok",ROE_wk!$AR$2:$AR$9915,Week_Overview!$AG$1,ROE_wk!$CI$2:$CI$9915,"<>RAO")'
        ws.Range("C7").Formula = '=COUNTIFS(ROE_wk!$A$2:$A$9915,"<>",ROE_wk!$AV$2:$AV$9915,"Ok",ROE_wk!$AR$2:$AR$9915,Week_Overview!$AG$1,ROE_wk!$CI$2:$CI$9915,"Sem Match")'
        ws.Range("E7").Formula = '=COUNTIFS(ROE_wk!$A$2:$A$9915,"<>",ROE_wk!$AV$2:$AV$9915,"Ok",ROE_wk!$AR$2:$AR$9915,Week_Overview!$AG$1,ROE_wk!$CI$2:$CI$9915,"<>RAO",ROE_wk!$CI$2:$CI$9915,"<>Sem Match")'

        headers = [
            "Nº OS",
            "Porto Final",
            "Region Final",
            "Match_Status",
            "Linha ROE_wk",
            "Booking",
            "Cliente",
            "Provedor",
            "Observação",
        ]
        for index, header in enumerate(headers, start=1):
            ws.Cells(9, index).Value = header

        ws.Range("J10:J250").Formula = '=IFERROR(AGGREGATE(15,6,ROW(ROE_wk!$A$2:$A$9915)/((ROE_wk!$A$2:$A$9915<>"")*(ROE_wk!$AV$2:$AV$9915="Ok")*(ROE_wk!$AR$2:$AR$9915=Week_Overview!$AG$1)*(ROE_wk!$CI$2:$CI$9915<>"RAO")),ROWS($J$10:J10)),"")'
        ws.Range("A10:A250").Formula = '=IF($J10="","",INDEX(ROE_wk!$A:$A,$J10))'
        ws.Range("B10:B250").Formula = '=IF($J10="","",INDEX(ROE_wk!$AY:$AY,$J10))'
        ws.Range("C10:C250").Formula = '=IF($J10="","",INDEX(ROE_wk!$AZ:$AZ,$J10))'
        ws.Range("D10:D250").Formula = '=IF($J10="","",INDEX(ROE_wk!$CI:$CI,$J10))'
        ws.Range("E10:E250").Formula = '=IF($J10="","",$J10)'
        ws.Range("F10:F250").Formula = '=IF($J10="","",INDEX(ROE_wk!$H:$H,$J10))'
        ws.Range("G10:G250").Formula = '=IF($J10="","",INDEX(ROE_wk!$M:$M,$J10))'
        ws.Range("H10:H250").Formula = '=IF($J10="","",INDEX(ROE_wk!$D:$D,$J10))'
        ws.Range("I10:I250").Formula = '=IF($J10="","",IF($D10="Sem Match","Sem RAO e sem inferência disponível","Sem RAO; usando inferência até a base oficial chegar"))'

        # Update ROE final formulas to read from the hidden static map instead of the visible dynamic area.
        table.ListColumns("Porto").DataBodyRange.Formula = '=IF(CG2<>"",CG2,IFERROR(XLOOKUP(A2,PortRegion_Fallback!L:L,PortRegion_Fallback!M:M,""),""))'
        table.ListColumns("Region").DataBodyRange.Formula = '=IF(CH2<>"",CH2,IFERROR(XLOOKUP(A2,PortRegion_Fallback!L:L,PortRegion_Fallback!N:N,""),""))'
        roe.Range(f"CI2:CI{last_roe_row}").Formula = '=IF(AND(CG2<>"",CH2<>""),"RAO",XLOOKUP(A2,PortRegion_Fallback!L:L,PortRegion_Fallback!O:O,"Sem Match"))'

        # Basic formatting to make the tab self-explanatory.
        ws.Range("A1:I1").Font.Bold = True
        ws.Range("A1:I1").Font.Size = 14
        ws.Range("A1:I1").Interior.Color = 12611584

        ws.Range("A2:I4").Font.Italic = True
        ws.Range("A2:I4").WrapText = True

        ws.Range("A6:E6").Font.Bold = True
        ws.Range("A6:E6").Interior.Color = 15921906
        ws.Range("A7:E7").Font.Bold = True

        ws.Range("A9:I9").Font.Bold = True
        ws.Range("A9:I9").Interior.Color = 14277081
        ws.Range("A9:I250").Borders.Weight = 2
        ws.Range("A:I").Columns.AutoFit()

        ws.Range("J:T").EntireColumn.Hidden = True
        workbook.Save()
        print("PortRegion_Fallback converted to dynamic view.")
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
