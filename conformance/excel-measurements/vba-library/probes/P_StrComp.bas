Attribute VB_Name = "P_StrComp"
Option Explicit
' StrComp (Strings). Result is a Variant Integer (-1, 0, 1) or Null.
Private Const N As Long = 25

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = StrComp("abc", "abc")                         ' equal
        Case 2: v = StrComp("abc", "abd")                         ' less
        Case 3: v = StrComp("abd", "abc")                         ' greater
        Case 4: v = StrComp("a", "A", vbBinaryCompare)            ' binary: a > A
        Case 5: v = StrComp("a", "A", vbTextCompare)              ' text: equal
        Case 6: v = StrComp("a", "A", vbDatabaseCompare)          ' database compare in Excel: error 5
        Case 7: v = StrComp("a", "A", 3)                          ' invalid compare: error 5
        Case 8: v = StrComp("a", "A", -1)                         ' default compare mode
        Case 9: v = StrComp(Null, "a")                            ' Null: Null
        Case 10: v = StrComp("a", Null)                           ' Null: Null
        Case 11: v = StrComp(Null, Null)                          ' both Null
        Case 12: v = StrComp(10, 9)                               ' numbers compare as strings
        Case 13: v = StrComp("", "")                              ' both empty
        Case 14: v = StrComp("", "a")                             ' empty vs text
        Case 15: v = StrComp(Empty, "")                           ' Empty vs empty
        Case 16: v = StrComp("a" & ChrW(228), "a" & ChrW(196), vbTextCompare)  ' non-ASCII text compare
        Case 17: v = StrComp("a" & ChrW(228), "az", vbTextCompare)  ' locale collation of a-umlaut vs z
        Case 18: v = StrComp("a" & ChrW(228), "az", vbBinaryCompare)  ' binary
        Case 19: v = StrComp("abc", "ab")                         ' prefix is smaller
        Case 20: v = TypeName(StrComp("a", "b"))                  ' result type
        Case 21: v = "compile-refused"
        Case 22: v = StrComp("a-b", "ab", vbTextCompare)          ' hyphen weighting; locale
        Case 23: v = StrComp("B", "a", vbTextCompare)             ' text ordering
        Case 24: v = StrComp("B", "a", vbBinaryCompare)           ' binary ordering
        Case 25: v = StrComp("a", "b", "1")                       ' numeric string compare mode
    End Select
    T = "StrComp." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "StrComp." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
