Attribute VB_Name = "P_Filter"
Option Explicit
' Filter (Strings). Returns a zero-based String() array; Include and Compare are optional.
Private Const N As Long = 24

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    Dim sa(1 To 3) As String, la(0 To 2) As Long, m(1 To 2, 1 To 2) As String
    Dim da() As String
    On Error GoTo fail
    sa(1) = "apple": sa(2) = "banana": sa(3) = "cherry"
    la(0) = 1: la(1) = 22: la(2) = 3
    m(1, 1) = "a": m(1, 2) = "b": m(2, 1) = "c": m(2, 2) = "d"
    Select Case k
        Case 1: v = Filter(Array("apple", "banana", "cherry"), "an")           ' include is the default
        Case 2: v = Filter(Array("apple", "banana", "cherry"), "an", False)    ' exclude
        Case 3: v = Filter(Array("apple", "banana", "cherry"), "AN")           ' binary: no match, bounds 0,-1
        Case 4: v = Filter(Array("apple", "banana", "cherry"), "AN", True, vbTextCompare)   ' text compare
        Case 5: v = Filter(Array("apple", "banana"), "")                       ' empty match string keeps all
        Case 6: v = Filter(Array(), "a")                                       ' empty source array
        Case 7: v = Filter(Array(1, 22, 3), "2")                               ' Variant array of numbers
        Case 8: v = Filter(la, "2")                                            ' Long array: error 13
        Case 9: v = Filter(Null, "a")                                          ' Null source
        Case 10: v = Filter("abc", "a")                                        ' not an array: error 13
        Case 11: v = Filter(Array("a", Null, "b"), "a")                        ' Null element
        Case 12: v = Filter(sa, "an")                                          ' 1-based String array: result is 0-based
        Case 13: v = Filter(Array("a", "b"), Null)                             ' Null match: error 94
        Case 14: v = Filter(SourceArray:=Array("x1", "y2"), Match:="2")        ' named arguments
        Case 15: v = Filter(Array("a1", "b2"), 2)                              ' numeric match
        Case 16: v = Filter(Array("a", "A"), "a", True, 1)                     ' compare as a number
        Case 17: v = Filter(m, "a")                                            ' 2-D array: error 5
        Case 18: v = Filter(Array("a", "b"), "a", Include:=False, Compare:=vbBinaryCompare)   ' named optional
        Case 19: v = Filter(Array("a", "b"), "a", True, 7)                     ' bad compare: error 5
        Case 20: v = Filter(da, "a")                                           ' unallocated dynamic array
        Case 21: v = Filter(Array("a", CVErr(5)), "a")                         ' error element
        Case 22: v = Filter(Array("ab", "AB", "aB"), "ab", True, vbTextCompare)   ' text compare keeps order
        Case 23: v = Filter(Array(ChrW(233), "E", "e"), "e", True, vbTextCompare) ' accent folding under text compare
        Case 24: v = Filter(Array("a", "b"), "a", Empty)                       ' Empty include = False
    End Select
    T = "Filter." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Filter." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
