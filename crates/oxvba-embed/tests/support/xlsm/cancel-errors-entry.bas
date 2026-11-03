Public Function RunProbe() As String
    On Error Resume Next
    ThisWorkbook.Save
    ThisWorkbook.Worksheets(1).Range("C1").Value2 = "caller resumed"
    RunProbe = "ok"
End Function
