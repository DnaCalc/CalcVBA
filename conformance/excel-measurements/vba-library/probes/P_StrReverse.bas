Attribute VB_Name = "P_StrReverse"
Option Explicit
' StrReverse (Strings).
Private Const N As Long = 14

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = StrReverse("abc")                 ' typical
        Case 2: v = StrReverse("")                    ' empty
        Case 3: v = StrReverse("a")                   ' single char
        Case 4: v = StrReverse(Null)                  ' Null: error 94
        Case 5: v = StrReverse(Empty)                 ' Empty
        Case 6: v = StrReverse(123)                   ' number as string argument
        Case 7: v = StrReverse("a" & ChrW(&HD83D) & ChrW(&HDE00))  ' surrogate pair: reversed per code unit?
        Case 8: v = StrReverse("e" & ChrW(&H301))     ' combining mark
        Case 9: v = StrReverse(" a ")                 ' spaces
        Case 10: v = StrReverse("ab" & vbCrLf)        ' CR LF
        Case 11: v = StrReverse(Expression:="xyz")    ' named argument
        Case 12: v = TypeName(StrReverse("ab"))       ' result type
        Case 13: v = StrReverse(1.5)                  ' double as string
        Case 14: v = StrReverse(ChrW(228) & "b")      ' non-ASCII
    End Select
    T = "StrReverse." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "StrReverse." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
