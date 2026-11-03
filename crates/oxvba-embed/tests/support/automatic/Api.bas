Public Journal As String
Public DocCount As Long
Public WatchCount As Long
Public CloseCount As Long
Public TermCount As Long
Public TermHasSource As Boolean
Public Nested As Boolean
Public observer As Watcher
Public Function ResetIdle() As Long
    DocCount = 0
    Journal = ""
    Nested = False
    ResetIdle = 0
End Function
Public Function ReadIdle() As String
    ReadIdle = DocCount & ":" & Journal
End Function
Public Function Scenario(ByVal book As Excel.Workbook) As String
    Dim sh As Excel.Worksheet
    Set sh = book.Worksheets(1)
    Set observer = New Watcher
    observer.Connect sh
    DocCount = 0
    WatchCount = 0
    Journal = ""
    sh.Range("A1").Value2 = 7#
    Scenario = DocCount & ":" & WatchCount & ":" & Journal
    Nested = True
    Journal = ""
    sh.Range("A1").Value2 = 8#
    Nested = False
    Scenario = Scenario & "|nested=" & DocCount & ":" & WatchCount & ":" & Journal
    observer.Connect sh
    Journal = ""
    sh.Range("A3").Value2 = 10#
    Scenario = Scenario & "|rebind=" & WatchCount & ":" & Journal
    observer.Disconnect
    sh.Range("A4").Value2 = 11#
    observer.Connect sh
    Set observer = Nothing
    Journal = ""
    sh.Range("A5").Value2 = 12#
    Scenario = Scenario & "|release=" & WatchCount & ":" & TermCount & ":" & TermHasSource & ":" & Journal
    book.Close SaveChanges:=False
    Scenario = Scenario & "|close=" & CloseCount
End Function
Public Function RunProbe() As String
    RunProbe = Scenario(ThisWorkbook)
End Function
