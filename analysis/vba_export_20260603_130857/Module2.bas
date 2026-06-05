Sub ROE_wk_monthly()
    Dim wbFonte As Workbook
    Dim wsFonte As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaLinhaFonte As Long, qtdeLinhas As Long
    Dim caminho As String
    Dim formulaColStart As Long, formulaColEnd As Long
    Dim col As Long
    Dim ultimaLinhaReal As Long, ultimaLinhaPlanilha As Long

    ' Otimização de performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Caminho do arquivo fonte
    caminho = "C:\Users\RMA704\OneDrive - Maersk Group\2025\Py\DSU 2025\ROE_wk_monthly\ROE_wk_monthly.xlsx"
    
    ' Abre o arquivo fonte como somente leitura
    Set wbFonte = Workbooks.Open(caminho, ReadOnly:=True)
    Set wsFonte = wbFonte.Sheets("Planilha1")
    Set wsDestino = ThisWorkbook.Sheets("ROE_wk_monthly")

    ' Identificar última linha com dados
    ultimaLinhaFonte = wsFonte.Cells(wsFonte.Rows.Count, "A").End(xlUp).Row
    qtdeLinhas = ultimaLinhaFonte - 5 ' dados começam em A6

    ' Limpar dados anteriores (A2:AK)
    wsDestino.Range("A2:AK" & wsDestino.Rows.Count).ClearContents

    ' Copiar dados da origem A6:AKx para A2 da planilha destino
    wsDestino.Range("A2").Resize(qtdeLinhas, 37).Value = _
        wsFonte.Range("A6").Resize(qtdeLinhas, 37).Value

    ' Fechar arquivo fonte
    wbFonte.Close SaveChanges:=False

    ' Copiar formatação da linha 2 para as linhas preenchidas
    wsDestino.Range("A2:AK2").Copy
    wsDestino.Range("A3:AK" & qtdeLinhas + 1).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' Aplicar fórmulas de AL2:BH2 nas novas linhas
    formulaColStart = wsDestino.Range("AL2").Column
    formulaColEnd = wsDestino.Range("BH2").Column

    ' Limpar fórmulas antigas (abaixo da linha 2)
    wsDestino.Range(wsDestino.Cells(3, formulaColStart), _
                    wsDestino.Cells(wsDestino.Rows.Count, formulaColEnd)).ClearContents

    ' Preencher fórmulas com base em AL2:BH2
    For col = formulaColStart To formulaColEnd
        If wsDestino.Cells(2, col).HasFormula Then
            wsDestino.Range(wsDestino.Cells(2, col), _
                            wsDestino.Cells(qtdeLinhas + 1, col)).Formula = _
                            wsDestino.Cells(2, col).Formula
        End If
    Next col

    ' Deletar linhas abaixo da última com conteúdo real
    ultimaLinhaReal = wsDestino.Cells(wsDestino.Rows.Count, "A").End(xlUp).Row
    ultimaLinhaPlanilha = wsDestino.Rows.Count
    If ultimaLinhaReal < ultimaLinhaPlanilha Then
        wsDestino.Range("A" & ultimaLinhaReal + 1 & ":A" & ultimaLinhaPlanilha).EntireRow.Delete
    End If

    ' Restaurar configurações
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub

