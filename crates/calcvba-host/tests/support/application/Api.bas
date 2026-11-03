Public Journal As String
Public AppCount As Long
Public DocCount As Long
Public SameBook As Boolean
Public observer As AppWatcher
Public Function RunProbe() As String
    Dim stage As Long
    On Error GoTo Failed
    stage = 1
    Set observer = New AppWatcher
    stage = 2
    observer.Connect ThisWorkbook.Application
    stage = 3
    ThisWorkbook.Close SaveChanges:=False
    RunProbe = AppCount & ":" & DocCount & ":" & Journal & ":" & SameBook
    Set observer = Nothing
    Exit Function
Failed:
    RunProbe = "FAIL:" & stage & ":" & Err.Number & ":" & Journal & ":" & AppCount & ":" & DocCount
End Function
