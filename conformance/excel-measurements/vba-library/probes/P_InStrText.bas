Attribute VB_Name = "P_InStrText"
Option Explicit
Option Compare Text
' Excel 16.0.20527.20196 follow-up probes.
Private Const N As Long = 34

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
        Case 13: v = InStr(1, "I", "i", 32768)
        Case 14: v = InStrRev("I", "i", -1, 32768)
        Case 15: v = InStr(1, "I", "", 32768)
        Case 16: v = InStrRev("I", "", -1, 32768)
        Case 17: v = InStr(1, "I", "i", 65536)
        Case 18: v = InStrRev("I", "i", -1, 65536)
        Case 19: v = InStr(1, "I", "", 65536)
        Case 20: v = InStrRev("I", "", -1, 65536)
        Case 21: v = InStr(1, "I", "i", 2147483647)
        Case 22: v = InStrRev("I", "i", -1, 2147483647)
        Case 23: v = InStr(1, "I", "", 2147483647)
        Case 24: v = InStrRev("I", "", -1, 2147483647)
        Case 25: v = InStr(1, "I", "", -2147483647)
        Case 26: v = InStr(1, "I", "", -2)
        Case 27: v = InStr(1, "I", "", -1)
        Case 28: v = InStr(1, "I", "", 2)
        Case 29: v = InStr(-1, "", "I", 0)
        Case 30: v = InStr(-1, "", "", 0)
        Case 31: v = InStr(0, "", "I", 0)
        Case 32: v = InStr(0, "", "", 0)
        Case 33: v = InStr(2147483647, "", "I", 0)
        Case 34: v = InStr(2147483647, "", "", 0)
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
