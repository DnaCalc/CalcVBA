Public Journal As String
Public Function RunProbe() As String
    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets(1)
    Journal = ""
    ws.Range("A1:A2,C1:C2").Value2 = 5#
    ThisWorkbook.Application.EnableEvents = False
    ws.Range("E1").Value2 = Journal
    ThisWorkbook.Application.EnableEvents = True
    RunProbe = "ok"
End Function
