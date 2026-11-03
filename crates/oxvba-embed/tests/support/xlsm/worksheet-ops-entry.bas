Public Function RunProbe() As String
    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets(1)
    ws.Range("A1").Value2 = 2#
    ws.Range("A2").Value2 = 3#
    ws.Range("A3").Value2 = 4#
    ws.Range("A2").Insert Shift:=xlDown
    ws.Range("B1").Formula = "=SUM(A1:A4)"
    ws.Range("A3").Delete Shift:=xlUp
    ws.Range("C1").Value2 = ws.Range("B1").Value2
    ws.Range("A1").ClearContents
    ws.Range("A2").Font.Bold = True
    RunProbe = "ok"
End Function
