Public Function RunProbe() As String
    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets(1)
    ws.Range("A1").Value2 = 4#
    ws.Range("A2").Formula = "=A1*3"
    ws.Range("A1:A2").NumberFormat = "0.00"
    ws.Range("A1:A2").Copy
    ws.Range("C1").PasteSpecial Paste:=xlPasteValues
    ws.Range("A1:A2").Copy Destination:=ws.Range("E1")
    ThisWorkbook.Application.CutCopyMode = False
    RunProbe = "ok"
End Function
