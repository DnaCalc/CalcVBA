Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    SaveCalls = SaveCalls + 1
    SaveUiSeen = SaveAsUI
    Cancel = True
End Sub
