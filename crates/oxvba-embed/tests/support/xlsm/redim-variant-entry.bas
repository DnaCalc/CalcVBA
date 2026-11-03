Public Function RunProbe() As String
    Dim ws As Object, data As Variant, copied As Variant
    Set ws = ThisWorkbook.Worksheets(1)
    ReDim data(1 To 2, -1 To 0)
    data(1, -1) = "a"
    data(1, 0) = 2#
    data(2, -1) = "b"
    data(2, 0) = 4#
    ReDim Preserve data(1 To 2, -1 To 1)
    data(1, 1) = 3#
    data(2, 1) = 6#
    ws.Range("A1:C2").Value = data
    copied = ws.Range("A1:C2").Value
    ReDim Preserve copied(1 To 2, 1 To 2)
    ws.Range("E1:F2").Value = copied
    On Error Resume Next
    ReDim Preserve data(1 To 3, -1 To 1)
    ws.Range("H1").Value2 = Err.Number
    On Error GoTo 0
    ws.Range("H2").Value2 = data(2, 1)
    RunProbe = "ok"
End Function
