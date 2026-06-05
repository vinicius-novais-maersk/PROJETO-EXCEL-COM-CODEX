Sub ROE_wk()
    Dim wbFonte As Workbook
    Dim wsFonte As Worksheet, wsDestino As Worksheet
    Dim ultimaLinhaFonte As Long, qtdeLinhas As Long
    Dim caminho As String
    Dim formulaColStart As Long, formulaColEnd As Long
    Dim col As Long

    ' Otimização
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.AskToUpdateLinks = False
    On Error GoTo FIM

    caminho = "C:\Users\RMA704\OneDrive - Maersk Group\2025\Py\DSU 2025\ROE_wk\ROE_wk.xlsx"
    Set wbFonte = Workbooks.Open(caminho, ReadOnly:=True, UpdateLinks:=0)
    Set wsFonte = wbFonte.Sheets("Planilha1")
    Set wsDestino = ThisWorkbook.Sheets("ROE_wk")

    ' Tamanho da origem (A6..)
    ultimaLinhaFonte = wsFonte.Cells(wsFonte.Rows.Count, "A").End(xlUp).Row
    qtdeLinhas = Application.Max(ultimaLinhaFonte - 5, 0)

    ' 1) LIMPA TUDO ABAIXO DA LINHA 2 (dados e fórmulas)
    '    - dados A:AK
    wsDestino.Range(wsDestino.Cells(3, "A"), wsDestino.Cells(wsDestino.Rows.Count, "AK")).ClearContents
    '    - fórmulas AL:BO
    wsDestino.Range(wsDestino.Cells(3, "AL"), wsDestino.Cells(wsDestino.Rows.Count, "BO")).ClearContents

    ' 2) CARREGA DADOS
    If qtdeLinhas > 0 Then
        wsDestino.Range("A2").Resize(qtdeLinhas, 37).Value = _
            wsFonte.Range("A6").Resize(qtdeLinhas, 37).Value
    End If

    wbFonte.Close SaveChanges:=False

    ' 3) FORMATAÇÃO (replica da linha 2)
    If qtdeLinhas > 0 Then
        wsDestino.Range("A2:AK2").Copy
        wsDestino.Range("A3:AK" & qtdeLinhas + 1).PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
    End If

    ' 4) PREENCHE FÓRMULAS (AL:BO) APENAS ATÉ AS NOVAS LINHAS
    formulaColStart = wsDestino.Range("AL2").Column
    formulaColEnd = wsDestino.Range("BO2").Column
    If qtdeLinhas > 0 Then
        For col = formulaColStart To formulaColEnd
            If wsDestino.Cells(2, col).HasFormula Then
                wsDestino.Range(wsDestino.Cells(2, col), _
                                wsDestino.Cells(qtdeLinhas + 1, col)).Formula = _
                                wsDestino.Cells(2, col).Formula
            End If
        Next col
    End If


FIM:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.AskToUpdateLinks = True
    If Err.Number <> 0 Then Err.Clear
End Sub

