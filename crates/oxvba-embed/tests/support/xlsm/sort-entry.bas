Public Function RunProbe() As String
    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets(1)
    ws.Range("A1").Value2 = "key"
    ws.Range("B1").Value2 = "value"
    ws.Range("A2").Value2 = 3#
    ws.Range("B2").Value2 = "c"
    ws.Range("A3").Value2 = 1#
    ws.Range("B3").Value2 = "a"
    ws.Range("A4").Value2 = 2#
    ws.Range("B4").Value2 = "b"
    ws.Range("A1:B4").Sort Key1:=ws.Range("A2"), Order1:=xlAscending, Header:=xlYes
    RunProbe = "ok"
End Function
