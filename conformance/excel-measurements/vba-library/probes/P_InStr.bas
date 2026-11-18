Attribute VB_Name = "P_InStr"
Option Explicit
' InStr (Strings). Module has no Option Compare, so compare -1 (vbUseCompareOption) means binary.
Private Const N As Long = 44

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = InStr("abcabc", "c")                              ' typical
        Case 2: v = InStr(4, "abcabc", "c")                           ' start inside
        Case 3: v = InStr(7, "abcabc", "c")                           ' start = Len+1
        Case 4: v = InStr(8, "abcabc", "c")                           ' start past the end
        Case 5: v = InStr(0, "abc", "a")                              ' start 0: error 5
        Case 6: v = InStr(-1, "abc", "a")                             ' start negative: error 5
        Case 7: v = InStr("abc", "")                                  ' empty search string: 1
        Case 8: v = InStr("", "a")                                    ' empty source
        Case 9: v = InStr("", "")                                     ' both empty
        Case 10: v = InStr(2, "abc", "")                              ' empty search with start
        Case 11: v = InStr(4, "abc", "")                              ' empty search, start = Len+1
        Case 12: v = InStr(5, "abc", "")                              ' empty search, start past end
        Case 13: v = InStr("abc", Null)                               ' Null search string
        Case 14: v = InStr(Null, "a")                                 ' Null source
        Case 15: v = InStr(Null, "abc", "a")                          ' Null start: error 94
        Case 16: v = InStr(1, "ABC", "b")                             ' binary: 0
        Case 17: v = InStr(1, "ABC", "b", vbTextCompare)              ' text: 2
        Case 18: v = InStr("ABC", "b", vbTextCompare)                 ' 3 args: first is Start, error 13
        Case 19: v = InStr(1, "ABC", "b", 0)                          ' compare as number 0 (binary)
        Case 20: v = InStr(1, "ABC", "b", 1)                          ' compare as number 1 (text)
        Case 21: v = InStr(1, "ABC", "b", 2)                          ' vbDatabaseCompare
        Case 22: v = InStr(1, "ABC", "b", 3)                          ' bad compare: error 5
        Case 23: v = InStr(1, "ABC", "b", -1)                         ' vbUseCompareOption: binary here
        Case 24: v = InStr(1, ChrW(233), "E", vbTextCompare)          ' e-acute vs E
        Case 25: v = InStr(1, ChrW(233), ChrW(201), vbTextCompare)    ' e-acute vs E-acute
        Case 26: v = InStr(1, ChrW(223), "ss", vbTextCompare)         ' sharp s vs ss
        Case 27: v = InStr(1, "ss", ChrW(223), vbTextCompare)         ' ss vs sharp s
        Case 28: v = InStr(1, ChrW(230), "ae", vbTextCompare)         ' ae ligature vs ae
        Case 29: v = InStr(1, "Stra" & ChrW(223) & "e", "STRASSE", vbTextCompare)   ' sharp s in a word
        Case 30: v = InStr(1, "xa" & ChrW(233), "e", vbTextCompare)   ' accent not folded to e
        Case 31: v = InStr(123456, 34)                                ' numbers as strings
        Case 32: v = InStr(1.5, "abc", "b")                           ' start rounds to 2
        Case 33: v = InStr("1", "abc", "b")                           ' numeric string start
        Case 34: v = "compile-refused"
        Case 35: v = "compile-refused"
        Case 36: v = "compile-refused"
        Case 37: v = InStr(1, Empty, "a")                             ' Empty source
        Case 38: v = InStr(1, "aaa", "aa")                            ' overlapping match
        Case 39: v = InStr(2, "aaa", "aa")                            ' overlapping, start 2
        Case 40: v = InStr(1, ChrW(26085) & ChrW(26412), ChrW(26412)) ' CJK
        Case 41: v = InStr(1, "x" & ChrW(55357) & ChrW(56832), ChrW(56832))   ' low surrogate found alone
        Case 42: v = InStr(1, "a" & ChrW(160) & "b", " ", vbTextCompare)      ' no-break space vs space
        Case 43: v = InStr(1, "I", "i", vbTextCompare)                ' simple case fold
        Case 44: v = InStr(2147483647, "abc", "a")                    ' huge start
    End Select
    T = "InStr." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "InStr." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
