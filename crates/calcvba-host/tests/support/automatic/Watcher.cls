Private WithEvents watched As Excel.Worksheet
Public Sub Connect(ByVal sh As Excel.Worksheet)
    Set watched = sh
End Sub
Public Sub Disconnect()
    Set watched = Nothing
End Sub
Private Sub watched_Change(ByVal Target As Excel.Range)
    WatchCount = WatchCount + 1
    Journal = Journal & "W" & Target.Address(False, False) & ";"
End Sub
Private Sub Class_Terminate()
    TermCount = TermCount + 1
    TermHasSource = Not (watched Is Nothing)
End Sub
