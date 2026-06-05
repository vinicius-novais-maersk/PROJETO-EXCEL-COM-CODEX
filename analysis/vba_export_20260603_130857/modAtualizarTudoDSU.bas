Option Explicit

Public Sub AtualizarTudoDSU(Optional ByVal showMessage As Boolean = True)
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldAskToUpdateLinks As Boolean
    Dim oldStatusBar As Variant
    Dim oldCalculation As XlCalculation
    Dim msg As String
    Dim etapa As String
    Dim falhou As Boolean

    On Error GoTo FalhaGeral

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldAskToUpdateLinks = Application.AskToUpdateLinks
    oldStatusBar = Application.StatusBar
    oldCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Application.Calculation = xlCalculationManual

    etapa = "Preparando conexoes"
    Application.StatusBar = "Atualizando DSU: " & etapa & "..."
    GarantirConexaoNoRefreshAll "Query - Sheet1"
    GarantirConexaoNoRefreshAll "Query - Reagendas"
    GarantirConexaoNoRefreshAll "Query - Reagendas_Filtro"
    GarantirConexaoNoRefreshAll "Query - Validacao_Final"

    etapa = "Atualizando Reagendas"
    AtualizarTabelaDSU "Reagendas", "Sheet1", "Query - Sheet1", etapa

    etapa = "Atualizando Reagendas_2"
    AtualizarTabelaDSU "Reagendas_2", "Reagendas", "Query - Reagendas", etapa

    etapa = "Atualizando Reagendas_Filtro"
    AtualizarTabelaDSU "Reagendas_Filtro", "Reagendas_Filtro", "Query - Reagendas_Filtro", etapa

    etapa = "Atualizando Validacao_Final"
    AtualizarTabelaDSU "Validacao_Final", "Validacao_Final", "Query - Validacao_Final", etapa

    etapa = "Sincronizando Validacao_BU"
    Application.StatusBar = "Atualizando DSU: " & etapa & "..."
    SincronizarValidacaoBU

    etapa = "Reatualizando Validacao_Final apos Validacao_BU"
    AtualizarTabelaDSU "Validacao_Final", "Validacao_Final", "Query - Validacao_Final", etapa

    etapa = "Recalculando workbook"
    Application.StatusBar = "Atualizando DSU: " & etapa & "..."
    Application.CalculateFull
    Application.CalculateUntilAsyncQueriesDone

    etapa = "Atualizando dashboards"
    On Error Resume Next
    AtualizarDashboardsSemPreenchimento False
    If Err.Number <> 0 Then
        msg = msg & vbCrLf & "Aviso na etapa '" & etapa & "': " & Err.Number & " - " & Err.Description
        Err.Clear
    End If
    On Error GoTo FalhaGeral

    etapa = "Salvando arquivo"
    ThisWorkbook.Save

Finalizar:
    Application.StatusBar = oldStatusBar
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    Application.AskToUpdateLinks = oldAskToUpdateLinks
    Application.Calculation = oldCalculation

    RegistrarResultadoAtualizacaoDSU falhou, msg, etapa

    If Not showMessage Then
        ' Modo silencioso para teste via automacao.
    ElseIf falhou Then
        MsgBox "Atualizacao DSU NAO foi concluida." & vbCrLf & msg, vbCritical, "Atualizar Tudo DSU"
    ElseIf Len(msg) > 0 Then
        MsgBox "Atualizacao DSU concluida com avisos:" & vbCrLf & msg, vbExclamation, "Atualizar Tudo DSU"
    Else
        MsgBox "Atualizacao DSU concluida." & vbCrLf & _
               "Bases, Validacao_Final, Validacao_BU e dashboards foram atualizados.", vbInformation, "Atualizar Tudo DSU"
    End If
    Exit Sub

FalhaGeral:
    falhou = True
    msg = msg & vbCrLf & "Etapa: " & etapa & vbCrLf & _
                "Erro: " & Err.Number & " - " & Err.Description
    Resume Finalizar
End Sub

Private Sub GarantirConexaoNoRefreshAll(ByVal connectionName As String)
    Dim cn As WorkbookConnection

    Set cn = ThisWorkbook.Connections(connectionName)
    cn.RefreshWithRefreshAll = True

    On Error Resume Next
    cn.OLEDBConnection.BackgroundQuery = False
    cn.ODBCConnection.BackgroundQuery = False
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub AtualizarTabelaDSU(ByVal sheetName As String, ByVal tableName As String, ByVal connectionName As String, ByVal friendlyName As String)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim cn As WorkbookConnection
    Dim tabelaErro As String
    Dim conexaoErro As String

    Application.StatusBar = "Atualizando DSU: " & friendlyName & "..."

    Set ws = ThisWorkbook.Worksheets(sheetName)
    Set lo = ws.ListObjects(tableName)

    On Error Resume Next
    Err.Clear
    lo.QueryTable.BackgroundQuery = False
    lo.QueryTable.Refresh BackgroundQuery:=False
    If Err.Number = 0 Then
        On Error GoTo 0
        Application.CalculateUntilAsyncQueriesDone
        Exit Sub
    End If
    tabelaErro = Err.Number & " - " & Err.Description
    Err.Clear
    On Error GoTo 0

    Set cn = ThisWorkbook.Connections(connectionName)
    On Error Resume Next
    Err.Clear
    cn.Refresh
    If Err.Number = 0 Then
        On Error GoTo 0
        Application.CalculateUntilAsyncQueriesDone
        Exit Sub
    End If
    conexaoErro = Err.Number & " - " & Err.Description
    Err.Clear
    On Error GoTo 0

    Err.Raise vbObjectError + 520, , friendlyName & " falhou. Tabela: " & tabelaErro & " | Conexao: " & conexaoErro
End Sub

Private Sub SincronizarValidacaoBU()
    Dim wsFinal As Worksheet
    Dim wsBU As Worksheet
    Dim tblFinal As ListObject
    Dim tblBU As ListObject
    Dim src As Range
    Dim linhasOrigem As Long
    Dim colunasOrigem As Long
    Dim linhasDestino As Long
    Dim colunasDestino As Long
    Dim destinoTabela As Range

    Set wsFinal = ThisWorkbook.Worksheets("Validacao_Final")
    Set wsBU = ThisWorkbook.Worksheets("Validacao_BU")
    Set tblFinal = wsFinal.ListObjects("Validacao_Final")
    Set tblBU = wsBU.ListObjects("Validacao_BU")

    If tblFinal.DataBodyRange Is Nothing Then Exit Sub

    Set src = tblFinal.DataBodyRange
    linhasOrigem = src.Rows.Count
    colunasOrigem = src.Columns.Count
    linhasDestino = tblBU.ListRows.Count
    colunasDestino = tblBU.ListColumns.Count

    If colunasDestino <> colunasOrigem Then
        Err.Raise vbObjectError + 521, , "Validacao_BU tem " & colunasDestino & " colunas, mas Validacao_Final tem " & colunasOrigem & "."
    End If

    If linhasDestino < linhasOrigem Then
        Set destinoTabela = tblBU.HeaderRowRange.Resize(linhasOrigem + 1, colunasOrigem)
        tblBU.Resize destinoTabela
    End If

    If Not tblBU.DataBodyRange Is Nothing Then
        tblBU.DataBodyRange.ClearContents
    End If

    tblBU.DataBodyRange.Cells(1, 1).Resize(linhasOrigem, colunasOrigem).Value = src.Value
End Sub


Private Sub RegistrarResultadoAtualizacaoDSU(ByVal falhou As Boolean, ByVal msg As String, ByVal etapa As String)
    Dim status As String

    If falhou Then
        status = "ERRO"
    ElseIf Len(msg) > 0 Then
        status = "AVISO"
    Else
        status = "OK"
    End If

    SetWorkbookNameText "DSU_LastUpdateStatus", status
    SetWorkbookNameText "DSU_LastUpdateEtapa", etapa
    SetWorkbookNameText "DSU_LastUpdateMessage", msg
    SetWorkbookNameText "DSU_LastUpdateAt", Format(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

Private Sub SetWorkbookNameText(ByVal nm As String, ByVal valueText As String)
    On Error Resume Next
    ThisWorkbook.Names(nm).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nm, RefersTo:="=""" & SafeNameText(valueText) & """"
End Sub

Private Function SafeNameText(ByVal valueText As String) As String
    SafeNameText = Replace(Replace(Left(CStr(valueText), 240), """", "'"), vbCrLf, " | ")
End Function
