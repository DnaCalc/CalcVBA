Attribute VB_Name = "P_InStrRev"
Option Explicit
' InStrRev (Strings). Parameters are StringCheck, StringMatch, Start, Compare; no Option Compare here.
Private Const N As Long = 28

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = InStrRev("abcabc", "c")                           ' typical
        Case 2: v = InStrRev("abcabc", "c", 5)                        ' start inside
        Case 3: v = InStrRev("abcabc", "c", -1)                       ' -1: from the end
        Case 4: v = InStrRev("abcabc", "c", 0)                        ' start 0: error 5
        Case 5: v = InStrRev("abcabc", "c", -2)                       ' start -2: error 5
        Case 6: v = InStrRev("abc", "c", 10)                          ' start past the end
        Case 7: v = InStrRev("abc", "")                               ' empty match
        Case 8: v = InStrRev("", "a")                                 ' empty source
        Case 9: v = InStrRev("", "")                                  ' both empty
        Case 10: v = InStrRev(Null, "a")                              ' Null source
        Case 11: v = InStrRev("abc", Null)                            ' Null match
        Case 12: v = InStrRev("abc", "c", Null)                       ' Null start
        Case 13: v = InStrRev("ABC", "b", -1, vbTextCompare)          ' text compare
        Case 14: v = InStrRev("ABC", "b", -1, vbBinaryCompare)        ' binary compare
        Case 15: v = InStrRev(ChrW(233), ChrW(201), -1, vbTextCompare)    ' e-acute vs E-acute
        Case 16: v = InStrRev(ChrW(233), "E", -1, vbTextCompare)      ' e-acute vs E
        Case 17: v = InStrRev(ChrW(223), "ss", -1, vbTextCompare)     ' sharp s vs ss
        Case 18: v = InStrRev(12321, 2)                               ' numbers as strings
        Case 19: v = InStrRev("abc", "c", "3")                        ' numeric string start
        Case 20: v = InStrRev(StringCheck:="abcb", StringMatch:="b")  ' named arguments
        Case 21: v = InStrRev("aaa", "aa")                            ' overlapping: 2
        Case 22: v = InStrRev("aaa", "aa", 2)                         ' overlapping, start 2
        Case 23: v = InStrRev("abc", "c", 3)                          ' start = Len
        Case 24: v = InStrRev("ABC", "b", -1, -1)                     ' vbUseCompareOption: binary here
        Case 25: v = InStrRev("ABC", "b", -1, 3)                      ' bad compare: error 5
        Case 26: v = InStrRev("ABC", "b", , vbTextCompare)            ' start omitted, compare given
        Case 27: v = InStrRev("abc", "", 2)                           ' empty match with start
        Case 28: v = InStrRev("abcabc", "b", Compare:=vbTextCompare, Start:=4)    ' named, out of order
    End Select
    T = "InStrRev." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "InStrRev." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
