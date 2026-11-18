Attribute VB_Name = "P_Left"
Option Explicit
' Left / Left$ (Strings; MS-VBAL 6.1.2.11.1.x).
Private Const N As Long = 24

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = Left("abcdef", 3)                ' typical
        Case 2: v = Left("abc", 0)                   ' zero length
        Case 3: v = Left("abc", 10)                  ' past the end
        Case 4: v = Left("abc", -1)                  ' negative length: error 5
        Case 5: v = Left(Null, 2)                    ' Null propagates
        Case 6: v = Left("abc", Null)                ' Null length: error 94
        Case 7: v = Left$("abc", 2)                  ' $ form: String
        Case 8: v = Left(String:="abc", Length:=2)   ' named arguments (VBE7 parameter names)
        Case 9: v = Left(12345, "2")                 ' numeric argument and numeric string
        Case 10: v = Left("abc", 3)                  ' length = Len
        Case 11: v = Left("", 1)                     ' empty string
        Case 12: v = Left(Empty, 2)                  ' Empty is ""
        Case 13: v = Left("abc", 1.5)                ' length rounds to even: 2
        Case 14: v = Left("abc", 0.5)                ' rounds to 0
        Case 15: v = Left("abc", "x")                ' non-numeric length: error 13
        Case 16: v = Left$(Null, 2)                  ' $ form with Null: error 94
        Case 17: v = Left("abc", 2147483647)         ' largest Long length
        Case 18: v = Left("abc", 2147483648#)        ' too large for Long: error 6
        Case 19: v = Left("abc", True)               ' True is -1: error 5
        Case 20: v = Left(ChrW(55357) & ChrW(56832), 1)   ' half a surrogate pair
        Case 21: v = Left(ChrW(160) & "ab", 2)       ' no-break space kept
        Case 22: v = Left("a" & ChrW(233) & ChrW(8364), 2)    ' non-ASCII characters
        Case 23: v = Left$(String:="abc", Length:=1) ' $ form, named
        Case 24: v = Left(vbNullString, 2)           ' vbNullString
    End Select
    T = "Left." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Left." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
