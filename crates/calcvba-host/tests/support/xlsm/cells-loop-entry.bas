Public Function RunProbe() As String
    Dim ws As Object, r As Long, c As Long, total As Long
    Set ws = ThisWorkbook.Worksheets(1)
    For r = 1 To 3
        For c = 1 To 4
            ws.Cells(r, c).Value2 = r * 10 + c
            total = total + ws.Cells(r, c).Value2
        Next c
    Next r
    ws.Range("F1").Value2 = total
    RunProbe = "ok"
End Function
