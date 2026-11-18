Attribute VB_Name = "P_InputBox"
Option Explicit
' InputBox (Interaction; never called, only constants and values relevant to it).
Private Const N As Long = 3

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Len(vbNullString)                                                     ' the cancelled-InputBox result is a zero-length string
        Case 2: v = vbUseDefault                                                          ' unused position
        Case 3: v = vbNullString = ""                                                     ' vbNullString equals the empty string
    End Select
    T = "InputBox." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "InputBox." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
