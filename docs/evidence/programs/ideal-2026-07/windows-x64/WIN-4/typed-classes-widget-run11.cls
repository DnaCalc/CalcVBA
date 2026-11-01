Option Explicit
Private Count As Long
Private Sub Class_Initialize()
    Count = 40
    Journal = Journal & "I"
End Sub
Public Function ReadCount() As Long
    ReadCount = Count
End Function
Public Sub Advance(ByVal amount As Long)
    Count = Count + amount
End Sub
Private Sub Class_Terminate()
    Journal = Journal & "T" & Count
End Sub
