Public Function RunProbe() As String
    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets(1)
    ws.Range("A1").Value2 = "group"
    ws.Range("A2").Value2 = "keep"
    ws.Range("A3").Value2 = "skip"
    ws.Range("A4").Value2 = "keep"
    ws.Range("A1:A4").AutoFilter Field:=1, Criteria1:="keep"
    ws.Range("C1").Value2 = ws.Rows(2).Hidden
    ws.Range("C2").Value2 = ws.Rows(3).Hidden
    ws.Range("C3").Value2 = ws.Rows(4).Hidden
    ws.ShowAllData
    ws.AutoFilterMode = False
    ws.Range("C4").Value2 = ws.Rows(3).Hidden
    RunProbe = "ok"
End Function
