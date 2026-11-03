Public Function RunProbe() As String
    Dim ws As Excel.Worksheet, cell As Excel.Range, total As Double
    Set ws = ThisWorkbook.Worksheets(1)
    ws.Range("A1").Value2 = 2#
    ws.Range("A2").Value2 = 3#
    ws.Range("B1").Value2 = 5#
    For Each cell In ws.UsedRange
        If cell.Value2 <> "" Then total = total + cell.Value2
    Next cell
    ws.Range("C1").Value2 = total
    RunProbe = "ok"
End Function
