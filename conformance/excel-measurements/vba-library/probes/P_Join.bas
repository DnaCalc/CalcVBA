Attribute VB_Name = "P_Join"
Option Explicit
' Join (Strings). Returns String; the source must be a one-dimensional array of String or Variant.
Private Const N As Long = 22

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    Dim sa(1 To 3) As String, la(0 To 2) As Long, m(1 To 2, 1 To 2) As String
    Dim da() As String
    On Error GoTo fail
    sa(1) = "x": sa(2) = "y": sa(3) = "z"
    la(0) = 1: la(1) = 2: la(2) = 3
    m(1, 1) = "a": m(1, 2) = "b": m(2, 1) = "c": m(2, 2) = "d"
    Select Case k
        Case 1: v = Join(Array("a", "b", "c"), ",")                    ' typical
        Case 2: v = Join(Array("a", "b", "c"))                         ' delimiter omitted: a space
        Case 3: v = Join(Array(), ",")                                 ' empty array
        Case 4: v = Join(Array(1, 2, 3), "-")                          ' numbers
        Case 5: v = Join(Array("a", Null, "b"), ",")                   ' Null element
        Case 6: v = Join(Array("a", Empty, "b"), ",")                  ' Empty element
        Case 7: v = Join("abc", ",")                                   ' not an array: error 13
        Case 8: v = Join(Null, ",")                                    ' Null source
        Case 9: v = Join(Array("a", "b"), "")                          ' empty delimiter
        Case 10: v = Join(Array("a", "b", "c"), "<>")                  ' multi-character delimiter
        Case 11: v = Join(Array("a", "b"), Null)                       ' Null delimiter
        Case 12: v = Join(Array("a", "b"), 0)                          ' numeric delimiter
        Case 13: v = Join(sa, "+")                                     ' 1-based String array
        Case 14: v = Join(Array(1.5, True, "x"), "|")                  ' mixed types ' locale
        Case 15: v = Join(Array("a", Array("b")), ",")                 ' nested array element: error 13
        Case 16: v = Join(SourceArray:=Array("p", "q"), Delimiter:=";")    ' named arguments
        Case 17: v = Join(m, ",")                                      ' 2-D array: error 5
        Case 18: v = Join(la, ",")                                     ' Long array
        Case 19: v = Join(Array("a", CVErr(5)), ",")                   ' error element
        Case 20: v = Join(da, ",")                                     ' unallocated dynamic array
        Case 21: v = Join(Array(ChrW(233), ChrW(8364)), ChrW(160))     ' non-ASCII elements and delimiter
        Case 22: v = Join(Array("a"), ",")                             ' single element: no delimiter
    End Select
    T = "Join." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Join." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
