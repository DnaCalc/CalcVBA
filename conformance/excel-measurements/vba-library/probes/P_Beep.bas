Attribute VB_Name = "P_Beep"
Option Explicit
' Beep (Interaction; a Sub that plays the default sound, called once).
Private Const N As Long = 3

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: Beep: v = "returned"                              ' one call
        Case 2: v = Err.Number                                    ' no error state left behind
        Case 3: v = "no further calls"                            ' Beep is called only once
    End Select
    T = "Beep." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Beep." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
