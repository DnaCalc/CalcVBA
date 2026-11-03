Public Function RunProbe() As String
    Dim ws As Object, items As Collection
    Set ws = ThisWorkbook.Worksheets(1)
    Set items = New Collection
    items.Add "red", "r"
    items.Add "blue", "b"
    items.Add "green", Before:=2
    ws.Range("A1").Value2 = items.Count
    ws.Range("A2").Value2 = items(2)
    ws.Range("A3").Value2 = items("b")
    items.Remove "r"
    ws.Range("A4").Value2 = items(1)
    RunProbe = "ok"
End Function
