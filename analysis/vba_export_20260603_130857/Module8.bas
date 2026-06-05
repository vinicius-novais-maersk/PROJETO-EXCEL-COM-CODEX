Sub ReplicarFormulas_ComDeslocamento()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Amb+RoFo+Cap")

    Dim rngOrigem As Range
    Dim rngDestino As Range

    Set rngOrigem = ws.Range("A16:BZ27") ' Pode ajustar colunas aqui se quiser (ex: "AR16:AX27")
    Set rngDestino = ws.Range("A2:BZ13") ' Mesma quantidade de linhas e colunas da origem

    Dim celOrigem As Range, celDestino As Range

    For Each celOrigem In rngOrigem
        Set celDestino = rngDestino.Cells(celOrigem.Row - rngOrigem.Row + 1, celOrigem.Column - rngOrigem.Column + 1)

        If celOrigem.HasFormula Then
            ' Copia e cola a fórmula com ajuste relativo, como se fosse arrastar
            celOrigem.Copy
            celDestino.PasteSpecial Paste:=xlPasteFormulas
        End If
    Next celOrigem

    Application.CutCopyMode = False
    MsgBox "Fórmulas replicadas com ajuste de referência!", vbInformation
End Sub
