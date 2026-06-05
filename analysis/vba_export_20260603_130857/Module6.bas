Sub SIL_wk()
    Dim wbFonte As Workbook
    Dim wsFonte As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaLinhaFonte As Long, ultimaColFonte As Long
    Dim caminho As String
    Dim qtdeLinhas As Long
    Dim colY As Long
    Dim ultimaLinhaDestino As Long
    Dim ultimaLinhaPlanilha As Long
    Dim nomePlanilhaFonte As String
    Dim diaAtual As String
    Dim sh As Worksheet

    On Error GoTo TratarErro

    ' Otimização de performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Caminho do arquivo fonte (ajuste se necessário)
    caminho = "C:\Users\RMA704\OneDrive - Maersk Group\2025\Py\DSU 2025\SIL_wk\SIL_wk_report"

    ' Dia do mês com dois dígitos
    diaAtual = Format(Day(Date), "00")

    ' Nome da aba dinâmica, ex.: PROGRAMAÇÕES DE TRANSPORTES_25-
    nomePlanilhaFonte = "PROGRAMAÇÕES DE TRANSPORTES_" & diaAtual & "-"

    ' Abre o arquivo fonte (tenta sem e com .xlsx)
    On Error Resume Next
    Set wbFonte = Workbooks.Open(caminho, ReadOnly:=True)
    If wbFonte Is Nothing Then
        Set wbFonte = Workbooks.Open(caminho & ".xlsx", ReadOnly:=True)
    End If
    On Error GoTo TratarErro

    If wbFonte Is Nothing Then
        MsgBox "Não foi possível abrir o arquivo fonte em:" & vbCrLf & caminho, vbCritical
        GoTo Finaliza
    End If

    ' Tenta localizar a aba exata; se não encontrar, tenta variações
    On Error Resume Next
    Set wsFonte = wbFonte.Sheets(nomePlanilhaFonte)                                ' com hífen final
    If wsFonte Is Nothing Then
        Set wsFonte = wbFonte.Sheets("PROGRAMAÇÕES DE TRANSPORTES_" & diaAtual)     ' sem hífen
    End If
    On Error GoTo TratarErro

    ' Última tentativa: qualquer aba que comece com o prefixo + dia
    If wsFonte Is Nothing Then
        For Each sh In wbFonte.Worksheets
            If sh.Name Like "PROGRAMAÇÕES DE TRANSPORTES_" & diaAtual & "*" Then
                Set wsFonte = sh
                Exit For
            End If
        Next sh
    End If

    If wsFonte Is Nothing Then
        MsgBox "Não encontrei uma aba no padrão 'PROGRAMAÇÕES DE TRANSPORTES_" & diaAtual & "-'.", vbCritical
        GoTo Finaliza
    End If

    ' Planilha de destino
    Set wsDestino = ThisWorkbook.Sheets("SIL_wk")

    ' Detectar última linha e coluna na fonte (a partir da linha 2)
    ultimaLinhaFonte = wsFonte.Cells(wsFonte.Rows.Count, "A").End(xlUp).Row
    ultimaColFonte = wsFonte.Cells(2, wsFonte.Columns.Count).End(xlToLeft).Column
    qtdeLinhas = Application.Max(0, ultimaLinhaFonte - 1) ' começa em A2

    ' Limpar conteúdo anterior no destino de A2 até a última coluna usada
    wsDestino.Range(wsDestino.Cells(2, 1), wsDestino.Cells(wsDestino.Rows.Count, ultimaColFonte)).ClearContents

    ' Copiar os valores da origem (A2:...) para A2 da planilha destino
    If qtdeLinhas > 0 Then
        wsDestino.Cells(2, 1).Resize(qtdeLinhas, ultimaColFonte).Value = _
            wsFonte.Cells(2, 1).Resize(qtdeLinhas, ultimaColFonte).Value
    End If

    ' Fechar o arquivo fonte
    wbFonte.Close SaveChanges:=False
    Set wbFonte = Nothing

    ' Copiar formatação da linha 2 para as novas linhas coladas
    If qtdeLinhas > 1 Then
        wsDestino.Rows(2).Copy
        wsDestino.Rows("3:" & qtdeLinhas + 1).PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
    End If

    ' Aplicar fórmula da coluna Y (a partir de Y2) até a última linha com dados
    colY = wsDestino.Range("Y2").Column

    ' Limpa fórmulas antigas abaixo da linha 2 na coluna Y
    wsDestino.Range(wsDestino.Cells(3, colY), _
                    wsDestino.Cells(wsDestino.Rows.Count, colY)).ClearContents

    ' Última linha com dados reais na coluna A
    ultimaLinhaDestino = wsDestino.Cells(wsDestino.Rows.Count, "A").End(xlUp).Row

    ' Arrasta a fórmula de Y2 para baixo, se existir
    If wsDestino.Cells(2, colY).HasFormula And ultimaLinhaDestino >= 2 Then
        wsDestino.Range(wsDestino.Cells(2, colY), _
                        wsDestino.Cells(ultimaLinhaDestino, colY)).Formula = _
                        wsDestino.Cells(2, colY).Formula
    End If

    ' Remover linhas abaixo da última com conteúdo
    ultimaLinhaPlanilha = wsDestino.Rows.Count
    If ultimaLinhaDestino < ultimaLinhaPlanilha Then
        wsDestino.Range("A" & ultimaLinhaDestino + 1 & ":A" & ultimaLinhaPlanilha).EntireRow.Delete
    End If

Finaliza:
    ' Restaurar configurações
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

TratarErro:
    MsgBox "Erro na macro SIL_wk:" & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
    On Error Resume Next
    If Not wbFonte Is Nothing Then wbFonte.Close SaveChanges:=False
    On Error GoTo 0
    Resume Finaliza
End Sub

