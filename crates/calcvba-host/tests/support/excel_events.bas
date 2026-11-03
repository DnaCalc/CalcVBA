Public eventCount As Long
Public beforeCloseCount As Long
Public suppressNested As Boolean
Public Sub Worksheet_Change(ByVal changedRange As Excel.Range)
    Dim app As Excel.Application
    Set app = changedRange.Application
    eventCount = eventCount + 1
    If changedRange.Address(False, False) = "F10" Then
        If suppressNested Then app.EnableEvents = False
        changedRange.Offset(0, 1).Value2 = 20#
        If suppressNested Then app.EnableEvents = True
    End If
End Sub
Public Sub Workbook_BeforeClose(ByRef Cancel As Boolean)
    beforeCloseCount = beforeCloseCount + 1
    Cancel = True
End Sub
Public Function StartChange(ByVal changedRange As Excel.Range, ByVal suppress As Boolean) As String
    Dim app As Excel.Application
    Set app = changedRange.Application
    eventCount = 0
    suppressNested = suppress
    changedRange.Value2 = 10#
    StartChange = eventCount & ":" & app.EnableEvents & ":" & changedRange.Offset(0, 1).Value2
End Function
Public Function TryClose(ByVal book As Excel.Workbook) As String
    beforeCloseCount = 0
    book.Close SaveChanges:=False
    TryClose = beforeCloseCount & ":" & book.Name
End Function
Public Function EventRange(ByVal sheet As Excel.Worksheet) As Excel.Range
    Set EventRange = sheet.Range("F10")
End Function
