Attribute VB_Name = "P_Replace"
Option Explicit
' Replace (Strings; MS-VBAL 6.1.2.11.1.x). Start > 1 drops the leading part of the result.
Private Const N As Long = 31

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = Replace("abcabc", "b", "X")                   ' typical
        Case 2: v = Replace("abcabc", "b", "X", 3)                ' start > 1 drops the leading part
        Case 3: v = Replace("aaaa", "a", "b", 1, 2)               ' count limits replacements
        Case 4: v = Replace("aaaa", "a", "b", 1, 0)               ' count 0
        Case 5: v = Replace("aaaa", "a", "b", 1, -1)              ' count -1: all
        Case 6: v = Replace("aaaa", "a", "b", 1, -2)              ' count -2: error 5 or all
        Case 7: v = Replace("abc", "b", "X", 0)                   ' start 0: error 5
        Case 8: v = Replace("abc", "b", "X", -1)                  ' start -1: error 5
        Case 9: v = Replace("abc", "b", "X", 4)                   ' start = Len+1: empty
        Case 10: v = Replace("abc", "b", "X", 10)                 ' start past the end: empty
        Case 11: v = Replace("ABC", "b", "X", , , vbTextCompare)  ' text compare
        Case 12: v = Replace("ABC", "b", "X", , , vbBinaryCompare) ' binary compare: no match
        Case 13: v = Replace("ABC", "b", "X", 1, -1, 2)           ' database compare in Excel
        Case 14: v = Replace("abc", "", "X")                      ' empty find string: unchanged
        Case 15: v = Replace("abc", "b", "")                      ' empty replacement deletes
        Case 16: v = Replace("abc", "b", "XYZ123")                ' replacement longer than find
        Case 17: v = Replace("aaa", "aa", "b")                    ' non-overlapping scan
        Case 18: v = Replace("", "a", "b")                        ' empty expression
        Case 19: v = Replace(Null, "a", "b")                      ' Null expression: error 94
        Case 20: v = Replace("a", Null, "b")                      ' Null find: error 94
        Case 21: v = Replace("a", "a", Null)                      ' Null replace: error 94
        Case 22: v = Replace(12345, 3, "x")                       ' numbers as arguments
        Case 23: v = TypeName(Replace("a", "a", "b"))             ' result type
        Case 24: v = Replace("a" & ChrW(196) & "b", ChrW(228), "X", 1, -1, vbTextCompare)  ' text compare, non-ASCII
        Case 25: v = Replace(Expression:="abcabc", Find:="b", Replace:="X", Start:=2, Count:=1, Compare:=vbBinaryCompare)  ' named
        Case 26: v = Replace("abcabc", "abc", "abc")              ' identical replacement
        Case 27: v = Replace("aaaa", "a", "b", 2, 2)              ' start and count together
        Case 28: v = Replace("abc", "abcd", "X")                  ' find longer than expression
        Case 29: v = Replace("abc", "b", "X", 2, 1, 5)            ' invalid compare: error 5
        Case 30: v = Replace("abc", "b", "X", "2")                ' numeric string start
        Case 31: v = Replace("a,b", ",", vbCrLf)                  ' replacement with control chars
    End Select
    T = "Replace." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Replace." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
