Option Explicit

Public Sub AtualizarDashboardsSemPreenchimento(Optional ByVal showMessage As Boolean = True)
    Dim targetSheets As Variant
    Dim sheetName As Variant
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim co As ChartObject
    Dim refreshedCaches As Object
    Dim cacheKey As String
    Dim activeWeek As Variant
    Dim updatedPivots As Long
    Dim updatedCharts As Long
    Dim updatedSheets As Long
    Dim updatedCaches As Long
    Dim resetFilters As Long
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldStatusBar As Variant
    Dim oldCalculation As XlCalculation
    Dim errors As String

    On Error GoTo FalhaGeral

    targetSheets = Array("Volume_DS", "Volume_MAO", "Volume_Graph", "Week_Overview", "Top_Offenders_Customers", "Top_Offenders_Vendors")
    Set refreshedCaches = CreateObject("Scripting.Dictionary")
    activeWeek = ThisWorkbook.Worksheets("Week_Overview").Range("AG1").Value

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldStatusBar = Application.StatusBar
    oldCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Atualizando dashboards DSU..."

    ' Atualiza o cache completo primeiro: isso nao depende do filtro visivel da pivot.
    For Each sheetName In targetSheets
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
        For Each pt In ws.PivotTables
            On Error Resume Next
            cacheKey = CStr(pt.CacheIndex)
            If Not refreshedCaches.Exists(cacheKey) Then
                pt.PivotCache.MissingItemsLimit = xlMissingItemsNone
                pt.PivotCache.Refresh
                If Err.Number <> 0 Then
                    errors = errors & vbCrLf & CStr(sheetName) & ": erro no cache da pivot '" & pt.Name & "' - " & Err.Description
                    Err.Clear
                Else
                    refreshedCaches.Add cacheKey, True
                    updatedCaches = updatedCaches + 1
                End If
            End If
            On Error GoTo FalhaGeral
        Next pt
    Next sheetName

    ' Atualiza visualizacoes. Limpa filtros temporarios, mas preserva filtros oficiais como OTO/Atrasado/Volume.
    For Each sheetName In targetSheets
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
        updatedSheets = updatedSheets + 1

        For Each pt In ws.PivotTables
            On Error Resume Next
            pt.ManualUpdate = True
            pt.PreserveFormatting = True
            resetFilters = resetFilters + LimparFiltrosTemporarios(pt, activeWeek)
            pt.ManualUpdate = False
            pt.RefreshTable
            If Err.Number <> 0 Then
                errors = errors & vbCrLf & CStr(sheetName) & ": erro na pivot '" & pt.Name & "' - " & Err.Description
                Err.Clear
                pt.ManualUpdate = False
            Else
                updatedPivots = updatedPivots + 1
            End If
            On Error GoTo FalhaGeral
        Next pt

        On Error Resume Next
        ws.Calculate
        If Err.Number <> 0 Then
            errors = errors & vbCrLf & CStr(sheetName) & ": erro ao recalcular aba - " & Err.Description
            Err.Clear
        End If
        On Error GoTo FalhaGeral

        For Each co In ws.ChartObjects
            On Error Resume Next
            co.Chart.Refresh
            If Err.Number <> 0 Then
                errors = errors & vbCrLf & CStr(sheetName) & ": erro no grafico '" & co.Name & "' - " & Err.Description
                Err.Clear
            Else
                updatedCharts = updatedCharts + 1
            End If
            On Error GoTo FalhaGeral
        Next co
    Next sheetName

    On Error Resume Next
    ThisWorkbook.Worksheets("Region errors").Calculate
    On Error GoTo FalhaGeral

    ThisWorkbook.Save

Finalizar:
    Application.StatusBar = oldStatusBar
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    Application.Calculation = oldCalculation

    If showMessage Then
        If Len(errors) > 0 Then
            MsgBox "Atualizacao concluida com avisos." & vbCrLf & _
                   "Abas: " & updatedSheets & vbCrLf & _
                   "Caches: " & updatedCaches & vbCrLf & _
                   "Pivots: " & updatedPivots & vbCrLf & _
                   "Graficos: " & updatedCharts & vbCrLf & _
                   "Filtros temporarios limpos: " & resetFilters & vbCrLf & vbCrLf & _
                   errors, vbExclamation, "Atualizar dashboards"
        Else
            MsgBox "Dashboards atualizados com sucesso." & vbCrLf & _
                   "Abas: " & updatedSheets & vbCrLf & _
                   "Caches: " & updatedCaches & vbCrLf & _
                   "Pivots: " & updatedPivots & vbCrLf & _
                   "Graficos: " & updatedCharts & vbCrLf & _
                   "Filtros temporarios limpos: " & resetFilters, vbInformation, "Atualizar dashboards"
        End If
    End If
    Exit Sub

FalhaGeral:
    errors = errors & vbCrLf & "Erro geral: " & Err.Description
    Resume Finalizar
End Sub

Private Function LimparFiltrosTemporarios(ByVal pt As PivotTable, ByVal activeWeek As Variant) As Long
    Dim pf As PivotField
    Dim n As String
    Dim resets As Long

    On Error Resume Next
    For Each pf In pt.PivotFields
        n = LCase$(pf.Name)
        If EhCampoSemana(n) Then
            pf.ClearAllFilters
            If pf.Orientation = xlPageField Then
                Err.Clear
                pf.CurrentPage = CStr(activeWeek)
                If Err.Number <> 0 Then
                    Err.Clear
                    pf.CurrentPage = CLng(activeWeek)
                End If
            End If
        ElseIf pf.Orientation = xlPageField And EhFiltroTemporario(n) Then
            pf.ClearAllFilters
            resets = resets + 1
        End If
    Next pf
    LimparFiltrosTemporarios = resets
End Function

Private Function EhCampoSemana(ByVal n As String) As Boolean
    EhCampoSemana = (InStr(1, n, "weeknum", vbTextCompare) > 0) _
        Or (InStr(1, n, "week num", vbTextCompare) > 0) _
        Or (InStr(1, n, "semana", vbTextCompare) > 0)
End Function

Private Function EhFiltroTemporario(ByVal n As String) As Boolean
    EhFiltroTemporario = (InStr(1, n, "region", vbTextCompare) > 0) _
        Or (InStr(1, n, "regiao", vbTextCompare) > 0) _
        Or (InStr(1, n, "porto", vbTextCompare) > 0) _
        Or (InStr(1, n, "port", vbTextCompare) > 0) _
        Or (InStr(1, n, "cliente", vbTextCompare) > 0) _
        Or (InStr(1, n, "customer", vbTextCompare) > 0) _
        Or (InStr(1, n, "provedor", vbTextCompare) > 0) _
        Or (InStr(1, n, "provider", vbTextCompare) > 0) _
        Or (InStr(1, n, "vendor", vbTextCompare) > 0)
End Function
