Private Sub Worksheet_Change(ByVal Target As Excel.Range)
    Dim part As Excel.Range
    Journal = Journal & Target.Areas.Count & ":" & Target.Count & ":"
    For Each part In Target.Areas
        Journal = Journal & part.Address(False, False) & ";"
    Next part
End Sub
