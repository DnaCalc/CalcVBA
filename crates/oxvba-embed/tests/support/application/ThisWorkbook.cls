Private Sub Workbook_BeforeClose(Cancel As Boolean)
    DocCount = DocCount + 1
    Journal = Journal & "D"
    Cancel = False
End Sub
