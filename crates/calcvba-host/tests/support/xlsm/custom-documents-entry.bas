Public Function RunProbe() As String
    Dim book As Object, ws As Object, boxed As Variant, n As Long
    Set book = ThisWorkbook
    Set ws = Sheet1
    Set boxed = Sheet1
    n = 3
    ws.Range("A1").Value2 = ws.Bump(n)
    ws.Range("A2").Value2 = n
    boxed.Caption = "custom"
    ws.Range("A3").Value2 = boxed.Caption
    ws.Range("A4").Value2 = book.Decorate("book")
    ws.Total = 3.5
    ws.Range("A5").Value2 = ws.Total
    ws.Range("B1").Value2 = (ws.Link Is Nothing)
    Set ws.Link = ws.Range("D1")
    ws.Range("B2").Value2 = (ws.Link Is ws.Range("D1"))
    Set ws.Link = Nothing
    ws.Range("B3").Value2 = (ws.Link Is Nothing)
    On Error Resume Next
    n = ws.SavedCaption
    ws.Range("A6").Value2 = Err.Number
    Err.Clear
    n = ws.Secret()
    ws.Range("A7").Value2 = Err.Number
    On Error GoTo 0
    RunProbe = "ok"
End Function
