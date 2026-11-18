Attribute VB_Name = "P_InStrB"
Option Explicit
' InStrB (Strings). Positions are byte positions (UTF-16LE: "abc" is 61 00 62 00 63 00).
Private Const N As Long = 18

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = InStrB("abc", "b")                        ' "b" is bytes 62 00, at byte 3
        Case 2: v = InStrB(1, "abc", ChrB(98))                ' single byte 62: byte 3
        Case 3: v = InStrB("abc", ChrB(0))                    ' the zero byte: byte 2
        Case 4: v = InStrB(2, "abc", "b")                     ' start at byte 2
        Case 5: v = InStrB(4, "abc", "b")                     ' start after the match
        Case 6: v = InStrB(0, "abc", "b")                     ' start 0: error 5
        Case 7: v = InStrB("abc", "")                         ' empty search: 1
        Case 8: v = InStrB("", "a")                           ' empty source
        Case 9: v = InStrB(Null, "a")                         ' Null source
        Case 10: v = InStrB("abc", Null)                      ' Null search
        Case 11: v = InStrB(1, "ABC", "b", vbTextCompare)     ' text compare
        Case 12: v = InStrB(1, ChrW(233), ChrB(233))          ' low byte of e-acute: byte 1
        Case 13: v = InStrB(1, ChrW(256), ChrB(1))            ' high byte of &H100: byte 2
        Case 14: v = InStrB(1, "a" & ChrW(233), ChrW(233))    ' whole character: byte 3
        Case 15: v = "compile-refused"
        Case 16: v = "compile-refused"
        Case 17: v = InStrB(1, "abc", ChrB(99) & ChrB(0))     ' two-byte match
        Case 18: v = InStrB(7, "abc", "")                     ' empty search, start = LenB+1
    End Select
    T = "InStrB." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "InStrB." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
