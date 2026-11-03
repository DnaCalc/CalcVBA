Public Hits As Long
Public Function RunProbe() As String
    Dim ws As Object, caught As Long
    Set ws = ThisWorkbook.Worksheets(1)
    On Error Resume Next
    ws.Range("A1").Value2 = 7#
    caught = Err.Number
    On Error GoTo 0
    ThisWorkbook.Application.EnableEvents = False
    ws.Range("C1").Value2 = Hits
    ws.Range("C2").Value2 = caught
    ThisWorkbook.Application.EnableEvents = True
    RunProbe = "ok"
End Function
