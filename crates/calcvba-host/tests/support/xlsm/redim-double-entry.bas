Public Function RunProbe() As String
    Dim ws As Object, data() As Double
    Set ws = ThisWorkbook.Worksheets(1)
    ReDim Preserve data(2 To 3, 4 To 5)
    data(2, 4) = 1.25
    data(2, 5) = 2.5
    data(3, 4) = 3.75
    data(3, 5) = 5#
    ReDim Preserve data(2 To 3, 4 To 6)
    ws.Range("A1:C2").Value = data
    ReDim data(1 To 1, 1 To 1)
    ws.Range("E1").Value2 = data(1, 1)
    RunProbe = "ok"
End Function
