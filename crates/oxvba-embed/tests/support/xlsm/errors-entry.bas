Public Function RunProbe() As String
    Dim ws As Object, value As Double
    Set ws = ThisWorkbook.Worksheets(1)
    On Error Resume Next
    value = 1# / 0#
    ws.Range("A1").Value2 = Err.Number
    Err.Clear
    ws.Range("A2").Value2 = Err.Number
    On Error GoTo Failed
    value = 1# / 0#
    ws.Range("A3").Value2 = "missed"
    RunProbe = "ok"
    Exit Function
Failed:
    ws.Range("A3").Value2 = Err.Number
    Resume Done
Done:
    ws.Range("A4").Value2 = "resumed"
    RunProbe = "ok"
End Function
