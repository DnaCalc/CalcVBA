Attribute VB_Name = "P_InStrText"
Option Explicit
Option Compare Text
' Excel 16.0.20527.20196 follow-up probes.
Private Const N As Long = 12

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = InStr(CLng(1073741825), "abc", "b")
        Case 2: v = InStr(CLng(2147483646), "abc", "b")
        Case 3: v = InStr(CLng(2147483647), "abc", "")
        Case 4: v = InStr(1, "ABC", "b", -2)
        Case 5: v = InStr(1, "ABC", "b", 4)
        Case 6: v = InStr(1, "ABC", "b", 1033)
        Case 7: v = InStr(1, "ABC", "b", 1055)
        Case 8: v = InStrRev("ABC", "b", -1, -2)
        Case 9: v = InStrRev("ABC", "b", -1, 2)
        Case 10: v = InStrRev("ABC", "b", -1, 4)
        Case 11: v = InStrRev("ABC", "b", -1, 1033)
        Case 12: v = InStrRev("ABC", "b", -1, 1055)
    End Select
    T = "InStrText." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "InStrText." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
