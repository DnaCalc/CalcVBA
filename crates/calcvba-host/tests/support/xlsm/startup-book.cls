Private Sub Workbook_Open()
    Opened = Opened + 1
    Me.Worksheets(1).Range("A1").Value2 = "opened"
End Sub
