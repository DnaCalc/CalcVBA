Private Sub Worksheet_Change(ByVal Target As Excel.Range)
    DocCount = DocCount + 1
    Journal = Journal & "D" & Target.Address(False, False) & ";"
    If Nested Then
        If Target.Address(False, False) = "A1" Then Target.Offset(1).Value2 = 9#
    End If
End Sub
