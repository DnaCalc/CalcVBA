Private Sub Workbook_BeforeClose(Cancel As Boolean)
    CloseCount = CloseCount + 1
    Cancel = True
End Sub
