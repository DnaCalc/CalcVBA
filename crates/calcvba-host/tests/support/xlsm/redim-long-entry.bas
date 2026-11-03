Public Function RunProbe() As String
    Dim ws As Excel.Worksheet, data() As Long, copied As Variant
    Set ws = ThisWorkbook.Worksheets(1)
    ReDim data(1 To 2, 1 To 2)
    data(1, 1) = 11
    data(1, 2) = 12
    data(2, 1) = 21
    data(2, 2) = 22
    ReDim Preserve data(1 To 2, 1 To 3)
    data(1, 3) = 13
    data(2, 3) = 23
    ws.Range("A1:C2").Value = data
    copied = ws.Range("A1:C2").Value
    ws.Range("E1:G2").Value = copied
    ws.Range("I1").Value2 = VarType(data)
    RunProbe = "ok"
End Function
