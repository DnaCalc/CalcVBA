Option Explicit
Public Journal As String
Public Function MakeWidget() As Widget
    Set MakeWidget = New Widget
End Function
Public Function ReadWidget(ByVal value As Widget) As Long
    ReadWidget = value.ReadCount()
End Function
Public Function RunProbe() As String
    Dim a As Widget
    Dim b As Widget
    Set a = New Widget
    Set b = a
    RunProbe = "same=" & (a Is b) & "|start=" & a.ReadCount()
    a.Advance 2
    RunProbe = RunProbe & "|next=" & b.ReadCount()
    Set b = New Widget
    RunProbe = RunProbe & "|distinct=" & (a Is b)
    Set a = Nothing
    Set b = Nothing
    RunProbe = RunProbe & "|journal=" & Journal
End Function
Public Function ReadJournal() As String
    ReadJournal = Journal
End Function
