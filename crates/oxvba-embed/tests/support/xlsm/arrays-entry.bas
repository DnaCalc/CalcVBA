Public Function RunProbe() As String
    Dim ws As Object, copied As Variant
    Dim data(1 To 2, 1 To 2) As Variant
    Set ws = ThisWorkbook.Worksheets(1)
    data(1, 1) = 1#
    data(1, 2) = 2#
    data(2, 1) = 3#
    data(2, 2) = 4#
    ws.Range("A1:B2").Value2 = data
    copied = ws.Range("A1:B2").Value2
    ws.Range("C1").Formula = "=SUM(A1:B2)"
    ws.Range("D1").Resize(2, 2).Value2 = copied
    RunProbe = copied(1, 1) & ":" & copied(2, 2) & ":" & ws.Range("C1").Value2
End Function
