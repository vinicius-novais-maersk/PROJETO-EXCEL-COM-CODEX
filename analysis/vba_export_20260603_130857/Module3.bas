Sub RAO_wk()
    Dim wbFonte As Workbook
    Dim wsFonte As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaLinhaFonte As Long, ultimaColFonte As Long
    Dim qtdeLinhas As Long
    Dim caminho As String
    Dim col As Long
    Dim ultimaLinhaDestino As Long
    Dim formulaColStart As Long, formulaColEnd As Long
    Dim ultimaLinhaPlanilha As Long

    ' Otimização de performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Caminho do arquivo fonte
    caminho = "C:\Users\RMA704\OneDrive - Maersk Group\2025\Py\DSU 2025\RAO_wk\RAO_wk.xlsx"
    
    ' Abre o arquivo fonte como somente leitura
    Set wbFonte = Workbooks.Open(caminho, ReadOnly:=True)
    Set wsFonte = wbFonte.Sheets("Planilha1")
    Set wsDestino = ThisWorkbook.Sheets("RAO_wk")

    ' Última linha e coluna com dados na origem
    ultimaLinhaFonte = wsFonte.Cells(wsFonte.Rows.Count, "A").End(xlUp).Row
    ultimaColFonte = wsFonte.Cells(6, wsFonte.Columns.Count).End(xlToLeft).Column
    qtdeLinhas = ultimaLinhaFonte - 5 ' dados começam em A6

    ' Limpar conteúdo anterior em A2 até a última coluna detectada
    wsDestino.Range(wsDestino.Cells(2, 1), wsDestino.Cells(wsDestino.Rows.Count, ultimaColFonte)).ClearContents

    ' Copiar valores de A6:Ex para A2
    wsDestino.Cells(2, 1).Resize(qtdeLinhas, ultimaColFonte).Value = _
        wsFonte.Cells(6, 1).Resize(qtdeLinhas, ultimaColFonte).Value

    ' Fechar o arquivo de origem
    wbFonte.Close SaveChanges:=False

    ' Aplicar formatação da linha 2 para as novas linhas
    wsDestino.Rows(2).Copy
    wsDestino.Rows(3 & ":" & qtdeLinhas + 1).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' Aplicar fórmulas de F2:G2
    formulaColStart = wsDestino.Range("F2").Column
    formulaColEnd = wsDestino.Range("G2").Column

    ' Limpar fórmulas antigas (abaixo da linha 2)
    wsDestino.Range(wsDestino.Cells(3, formulaColStart), _
                    wsDestino.Cells(wsDestino.Rows.Count, formulaColEnd)).ClearContents

    ' Última linha colada com base na coluna A
    ultimaLinhaDestino = wsDestino.Cells(wsDestino.Rows.Count, "A").End(xlUp).Row

    ' Aplicar as fórmulas de F2:G2 até a última linha colada
    For col = formulaColStart To formulaColEnd
        If wsDestino.Cells(2, col).HasFormula Then
            wsDestino.Range(wsDestino.Cells(2, col), _
                            wsDestino.Cells(ultimaLinhaDestino, col)).Formula = _
                            wsDestino.Cells(2, col).Formula
        End If
    Next col

    ' Deletar linhas extras após os dados
    ultimaLinhaPlanilha = wsDestino.Rows.Count
    If ultimaLinhaDestino < ultimaLinhaPlanilha Then
        wsDestino.Range("A" & ultimaLinhaDestino + 1 & ":A" & ultimaLinhaPlanilha).EntireRow.Delete
    End If

    ' Restaurar configurações
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub



