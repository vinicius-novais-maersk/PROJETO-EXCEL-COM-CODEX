Sub SIL_wk_month()
    Dim wbFonte As Workbook
    Dim wsFonte As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaLinhaFonte As Long, ultimaColFonte As Long
    Dim caminho As String
    Dim qtdeLinhas As Long
    Dim colY As Long, colZ As Long
    Dim ultimaLinhaDestino As Long
    Dim ultimaLinhaPlanilha As Long
    Dim nomePlanilhaFonte As String
    Dim diaAtual As String

    ' Otimização de performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Caminho do arquivo fonte
    caminho = "C:\Users\RMA704\OneDrive - Maersk Group\2025\Py\DSU 2025\SIL_wk\SIL_wk_month.xlsx"

    ' Obter o dia do mês com dois dígitos
    diaAtual = Format(Day(Date), "00")

    ' Montar o nome da aba dinamicamente
    nomePlanilhaFonte = "PROGRAMAÇÕES DE TRANSPORTES_" & diaAtual & "-"

    ' Abre o arquivo fonte
    Set wbFonte = Workbooks.Open(caminho, ReadOnly:=True)
    Set wsFonte = wbFonte.Sheets(nomePlanilhaFonte)
    Set wsDestino = ThisWorkbook.Sheets("SIL_wk_month")

    ' Detectar última linha e coluna na aba fonte (a partir da linha 2)
    ultimaLinhaFonte = wsFonte.Cells(wsFonte.Rows.Count, "A").End(xlUp).Row
    ultimaColFonte = wsFonte.Cells(2, wsFonte.Columns.Count).End(xlToLeft).Column
    qtdeLinhas = ultimaLinhaFonte - 1 ' pois começa em A2

    ' Limpar conteúdo anterior no destino de A2 até última coluna da origem
    wsDestino.Range(wsDestino.Cells(2, 1), wsDestino.Cells(wsDestino.Rows.Count, ultimaColFonte)).ClearContents

    ' Copiar os valores da origem (A2:Xx) para A2 da planilha destino
    wsDestino.Cells(2, 1).Resize(qtdeLinhas, ultimaColFonte).Value = _
        wsFonte.Cells(2, 1).Resize(qtdeLinhas, ultimaColFonte).Value

    ' Fechar o arquivo fonte
    wbFonte.Close SaveChanges:=False

    ' Copiar formatação da linha 2 para as novas linhas coladas
    wsDestino.Rows(2).Copy
    wsDestino.Rows("3:" & qtdeLinhas + 1).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' Colunas Y e Z
    colY = wsDestino.Range("Y2").Column
    colZ = wsDestino.Range("Z2").Column

    ' Limpa fórmulas antigas abaixo da linha 2 nas colunas Y e Z
    wsDestino.Range(wsDestino.Cells(3, colY), wsDestino.Cells(wsDestino.Rows.Count, colY)).ClearContents
    wsDestino.Range(wsDestino.Cells(3, colZ), wsDestino.Cells(wsDestino.Rows.Count, colZ)).ClearContents

    ' Última linha com dados reais na coluna A
    ultimaLinhaDestino = wsDestino.Cells(wsDestino.Rows.Count, "A").End(xlUp).Row

    ' Reaplica fórmula da linha 2 nas colunas Y e Z até a última linha preenchida
    If wsDestino.Cells(2, colY).HasFormula Then
        wsDestino.Range(wsDestino.Cells(2, colY), wsDestino.Cells(ultimaLinhaDestino, colY)).Formula = _
            wsDestino.Cells(2, colY).Formula
    End If

    If wsDestino.Cells(2, colZ).HasFormula Then
        wsDestino.Range(wsDestino.Cells(2, colZ), wsDestino.Cells(ultimaLinhaDestino, colZ)).Formula = _
            wsDestino.Cells(2, colZ).Formula
    End If

    ' Remover linhas abaixo da última com conteúdo
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

