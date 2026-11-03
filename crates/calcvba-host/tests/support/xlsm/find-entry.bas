Public Function RunProbe() As String
    Dim ws As Excel.Worksheet, found As Excel.Range
    Set ws = ThisWorkbook.Worksheets(1)
    ws.Range("A1").Value2 = "first"
    ws.Range("A2").Value2 = "needle"
    ws.Range("A3").Value2 = "last"
    Set found = ws.Range("A1:A3").Find(What:="needle", LookIn:=xlValues, LookAt:=xlWhole)
    ws.Range("B1").Value2 = found.Address(False, False)
    Set found = ws.Range("A1:A3").Find(What:="absent", LookIn:=xlValues, LookAt:=xlWhole)
    ws.Range("B2").Value2 = found Is Nothing
    RunProbe = "ok"
End Function
