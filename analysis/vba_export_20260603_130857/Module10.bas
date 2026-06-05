Sub AtualizarValidacaoBU()
    Dim wsFinal As Worksheet
    Dim wsBU As Worksheet
    Dim tblFinal As ListObject
    Dim tblBU As ListObject
    
    ' Definir as planilhas
    Set wsFinal = ThisWorkbook.Sheets("Validacao_Final") ' Nome da aba com a query final
    Set wsBU = ThisWorkbook.Sheets("Validacao_BU") ' Nome da aba onde está a tabela manual
    
    ' Definir as tabelas
    Set tblFinal = wsFinal.ListObjects(1) ' assume que a primeira tabela da aba é a correta
    Set tblBU = wsBU.ListObjects("Validacao_BU") ' nome exato da tabela manual

    ' Limpa os dados da tabela BU (exceto cabeçalho)
    tblBU.DataBodyRange.ClearContents

    ' Copia os dados da tabela Final
    tblFinal.DataBodyRange.Copy

    ' Cola na tabela BU (início do corpo da tabela)
    tblBU.DataBodyRange.Cells(1, 1).PasteSpecial Paste:=xlPasteValues

    MsgBox "Tabela Validacao_BU atualizada com sucesso!", vbInformation
End Sub
