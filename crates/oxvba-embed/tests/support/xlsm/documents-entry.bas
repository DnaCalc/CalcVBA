Public Function RunProbe() As String
    Dim book As Object, ws As Object, boxed As Variant
    Set book = ThisWorkbook
    Set boxed = Sheet1
    Set ws = book.Worksheets(1)
    ws.Range("A1").Value2 = 23#
    boxed.Range("B1").Value2 = ws.Range("A1").Offset(0, 1).Column
    RunProbe = CStr(book Is ThisWorkbook) & ":" & CStr(ws Is Sheet1)
    RunProbe = RunProbe & ":" & CStr(boxed Is ws) & ":" & CStr(ws.Parent Is ThisWorkbook)
    RunProbe = RunProbe & ":" & CStr(boxed.Range("B1").Value2)
End Function
