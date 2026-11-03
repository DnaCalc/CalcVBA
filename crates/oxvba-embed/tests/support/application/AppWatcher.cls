Private WithEvents watched As Excel.Application
Public Sub Connect(ByVal app As Excel.Application)
    Set watched = app
End Sub
Private Sub watched_WorkbookBeforeClose(ByVal Wb As Excel.Workbook, Cancel As Boolean)
    AppCount = AppCount + 1
    Journal = Journal & "A"
    SameBook = Wb Is ThisWorkbook
    Cancel = True
End Sub
