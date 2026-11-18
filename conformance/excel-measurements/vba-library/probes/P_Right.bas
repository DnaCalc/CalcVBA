Attribute VB_Name = "P_Right"
Option Explicit
' Right / Right$ (Strings).
Private Const N As Long = 15

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = Right("abcdef", 3)                ' typical
        Case 2: v = Right("abc", 0)                   ' zero length
        Case 3: v = Right("abc", 10)                  ' past the start
        Case 4: v = Right("abc", -1)                  ' negative length: error 5
        Case 5: v = Right(Null, 2)                    ' Null propagates
        Case 6: v = Right("abc", Null)                ' Null length: error 94
        Case 7: v = Right$("abc", 2)                  ' $ form: String
        Case 8: v = Right(String:="abc", Length:=2)   ' named arguments
        Case 9: v = Right(12345, "2")                 ' numeric argument and numeric string
        Case 10: v = Right("", 2)                     ' empty string
        Case 11: v = Right(Empty, 2)                  ' Empty
        Case 12: v = Right$(Null, 2)                  ' $ form with Null: error 94
        Case 13: v = Right("abc", 1.5)                ' fractional length rounds (banker's)
        Case 14: v = Right("abc", 2147483647)         ' largest Long length
        Case 15: v = Right$(String:="abc", Length:=1) ' named, $ form
    End Select
    T = "Right." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Right." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
