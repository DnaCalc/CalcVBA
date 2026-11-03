Public Function RunProbe() As String
    Dim ws As Object, a As Variant, b As Variant, s As String, n As Long, result As Variant, errorCode As Long, d As Double
    Set ws = ThisWorkbook.Worksheets(1)
    a = 2#: s = ""
    ws.Range("A1").Value2 = a <> s
    ws.Range("A2").Value2 = s < a
    a = 5&: b = "4"
    ws.Range("A3").Value2 = a < b
    ws.Range("A4").Value2 = a = "5"
    n = 5
    ws.Range("A5").Value2 = n > b
    a = Empty: b = "0"
    ws.Range("A6").Value2 = a = b
    a = Null: result = (a = s)
    ws.Range("A7").Value2 = VarType(result)
    a = "abc": b = 5
    ws.Range("B1").Value2 = (a = b)
    a = "5": b = 5
    ws.Range("B2").Value2 = (a < b)
    NumericCases ws
    OrderingCases ws
    RunProbe = "ok"
End Function

Private Sub NumericCases(ByVal ws As Object)
    Dim a As Variant, n As Long, result As Variant, errorCode As Long
    a = "abc": n = 5: result = Null
    On Error Resume Next
    result = (n = a)
    errorCode = Err.Number
    ws.Range("B3").Value2 = errorCode
    ws.Range("B4").Value2 = VarType(result)
    ws.Range("B5").Value2 = VarType(n)
    ws.Range("B6").Value2 = VarType(a)
    Err.Clear
    n = 12: a = "4"
    result = (n > a)
    errorCode = Err.Number
    ws.Range("B7").Value2 = result
    ws.Range("B8").Value2 = errorCode
    On Error GoTo 0
End Sub

Private Sub OrderingCases(ByVal ws As Object)
    Dim a As Variant, n As Long, d As Double, result As Variant, errorCode As Long
    On Error Resume Next
    Err.Clear
    n = 5: a = "!": result = Empty
    result = (n < a)
    errorCode = Err.Number
    ws.Range("B9").Value2 = errorCode
    ws.Range("B10").Value2 = result
    Err.Clear: result = Empty
    result = (n > a)
    errorCode = Err.Number
    ws.Range("B11").Value2 = errorCode
    ws.Range("B12").Value2 = result
    Err.Clear: result = Empty
    result = (n <> a)
    errorCode = Err.Number
    ws.Range("B13").Value2 = errorCode
    ws.Range("B14").Value2 = result
    Err.Clear: result = Empty
    d = 5#: a = "!"
    result = (d < a)
    errorCode = Err.Number
    ws.Range("B15").Value2 = result
    ws.Range("B16").Value2 = errorCode
    Err.Clear
    n = a
    errorCode = Err.Number
    ws.Range("B17").Value2 = errorCode
    On Error GoTo 0
End Sub
