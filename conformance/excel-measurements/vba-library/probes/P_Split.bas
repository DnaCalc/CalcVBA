Attribute VB_Name = "P_Split"
Option Explicit
' Split (Strings; MS-VBAL 6.1.2.11.1.x). Result is a zero-based String() in a Variant.
Private Const N As Long = 33

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = Split("a,b,c", ",")               ' typical
        Case 2: v = Split("", ",")                    ' empty string: UBound -1
        Case 3: v = UBound(Split("", ","))            ' UBound of the empty result
        Case 4: v = Split("a b c")                    ' default delimiter is space
        Case 5: v = Split("a,b,c", ",", 2)            ' limit
        Case 6: v = Split("a,b,c", ",", 1)            ' limit 1: whole string
        Case 7: v = Split("a,b,c", ",", 0)            ' limit 0
        Case 8: v = Split("a,b,c", ",", -1)           ' limit -1: all
        Case 9: v = Split("a,b,c", ",", -2)           ' limit -2: error 5 or all
        Case 10: v = Split("a,b,c", ",", 10)          ' limit above count
        Case 11: v = Split("a--b--c", "--")           ' multi-character delimiter
        Case 12: v = Split("abc", "")                 ' delimiter "": single element
        Case 13: v = Split("aXbxc", "x", -1, vbTextCompare)    ' text compare
        Case 14: v = Split("aXbxc", "x", -1, vbBinaryCompare)  ' binary compare
        Case 15: v = TypeName(Split("a,b", ","))      ' result type: String()
        Case 16: v = Split(Null, ",")                 ' Null expression: error 94
        Case 17: v = Split("a,b", Null)               ' Null delimiter
        Case 18: v = Split("a,,b", ",")               ' empty middle element
        Case 19: v = Split(",a,", ",")                ' empty first and last
        Case 20: v = Split(12321, 3)                  ' numbers as arguments
        Case 21: v = Split("abc", "abcd")             ' delimiter longer than string
        Case 22: v = Split("abc", "abc")              ' delimiter equals the string
        Case 23: v = Split(Expression:="a;b", Delimiter:=";", Limit:=-1, Compare:=vbBinaryCompare)  ' named
        Case 24: v = LBound(Split("a,b", ","))        ' lower bound is 0
        Case 25: v = Split("a,b,c", ",")(1)           ' index the result directly
        Case 26: v = Split(vbNullString)              ' vbNullString
        Case 27: v = Split(Empty, ",")                ' Empty expression
        Case 28: v = Split("a,b", ",", "2")           ' numeric string limit
        Case 29: v = Split("a" & vbTab & "b" & vbLf & "c", vbLf)  ' control-character delimiter
        Case 30: v = Split("a,b", ",", 2, 2)          ' database compare in Excel
        Case 31: v = Split("a,b", ",", 2, 5)          ' invalid compare: error 5
        Case 32: v = Split("aaa", "aa")               ' overlapping delimiter
        Case 33: v = Split("a b", Empty)              ' Empty delimiter
    End Select
    T = "Split." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Split." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
