Public Function RunProbe() As String
    Dim ws As Object, data As Payload, copied As Variant
    Set ws = ThisWorkbook.Worksheets(1)
    Set data = New Payload
    On Error Resume Next
    ws.Range("I1").Value2 = data.GetItem(1, 1)
    ws.Range("I2").Value2 = Err.Number
    On Error GoTo 0
    data.Fill
    data.PutItem 2, 3, 29
    ws.Range("A1:C2").Value2 = data.Matrix
    copied = ws.Range("A1:C2").Value2
    data.Values = copied
    ws.Range("E1:G2").Value2 = data.Values
    RunProbe = "ok"
End Function
