Public Function RunProbe() As String
    Dim ws As Object, typed() As Long, boxed As Variant, caught As Long
    Set ws = ThisWorkbook.Worksheets(1)
    ReDim typed(1 To 2)
    typed(1) = 5
    typed(2) = 6
    boxed = typed
    On Error Resume Next
    ReDim Preserve boxed(1 To 3)
    caught = Err.Number
    Err.Clear
    ws.Range("A1").Value2 = caught
    ws.Range("A2").Value2 = VarType(boxed)
    ws.Range("A3").Value2 = boxed(2)
    ws.Range("A4").Value2 = UBound(boxed)
    ReDim boxed(1 To 4)
    ws.Range("B1").Value2 = Err.Number
    Err.Clear
    ws.Range("B2").Value2 = VarType(boxed)
    ws.Range("B3").Value2 = boxed(2)
    ReDim boxed(1 To 2) As Double
    ws.Range("C1").Value2 = VarType(boxed)
    ReDim Preserve boxed(1 To 3) As Variant
    ws.Range("C2").Value2 = Err.Number
    ws.Range("C3").Value2 = VarType(boxed)
    On Error GoTo 0
    RunProbe = "ok"
End Function
