Public Function RunProbe() As String
    Dim ws As Object, text As String, i As Long
    Set ws = ThisWorkbook.Worksheets(1)
    For i = 1 To 4
        If i > 1 Then text = text & ","
        text = text & "item" & CStr(i)
    Next i
    ws.Range("A1").Value2 = text
    ws.Range("A2").Value2 = Left$(text, 5)
    ws.Range("A3").Value2 = Replace(text, "item", "row")
    ws.Range("A4").Value2 = UCase$(Trim$("  mixed Case  "))
    RunProbe = "ok"
End Function
