Option Explicit
Public Function RunProbe() As String
    Dim ws As Excel.Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    ws.Range("A1").Value2 = xlValues
    ws.Range("A2").Value2 = Excel.xlWhole
    ws.Range("A3").Value2 = Excel.XlPasteType.xlPasteValues
    ws.Range("A4").Value2 = xlAscending
    ws.Range("A5").Value2 = xlYes
    ws.Range("A6").Value2 = xlDown
    ws.Range("A7").Value2 = xlUp
    RunProbe = "ok"
End Function
