Public SaveCalls As Long
Public SaveUiSeen As Boolean
Public Function RunProbe() As String
    Dim book As Object, ws As Object
    Set book = ThisWorkbook
    Set ws = book.Worksheets(1)
    book.Save
    ws.Range("A1").Value2 = SaveCalls
    ws.Range("A2").Value2 = SaveUiSeen
    RunProbe = "ok"
End Function
