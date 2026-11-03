Public Opened As Long
Public Function RunProbe() As String
    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets(1)
    ws.Range("B1").Value2 = Opened
    ws.Range("B2").Value2 = ws.Range("A1").Value2
    RunProbe = "ok"
End Function
