Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim result As Double
    Cancel = True
    Me.Worksheets(1).Range("A1").Value2 = "before error"
    result = 1# / 0#
End Sub
