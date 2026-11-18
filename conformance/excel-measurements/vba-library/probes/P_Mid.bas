Attribute VB_Name = "P_Mid"
Option Explicit
' Mid / Mid$ (Strings; MS-VBAL 6.1.2.11.1.x).
Private Const N As Long = 23

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = Mid("abcdef", 2, 3)               ' typical
        Case 2: v = Mid("abcdef", 2)                  ' length omitted: to the end
        Case 3: v = Mid("abc", 4)                     ' start = Len+1: empty
        Case 4: v = Mid("abc", 5, 1)                  ' start past the end: empty
        Case 5: v = Mid("abc", 0)                     ' start 0: error 5
        Case 6: v = Mid("abc", -1, 1)                 ' negative start: error 5
        Case 7: v = Mid("abc", 2, -1)                 ' negative length: error 5
        Case 8: v = Mid("abc", 2, 0)                  ' zero length
        Case 9: v = Mid("abc", 2, 100)                ' length past the end
        Case 10: v = Mid(Null, 1, 1)                  ' Null propagates
        Case 11: v = Mid("abc", Null, 1)              ' Null start: error 94
        Case 12: v = Mid("abc", 1, Null)              ' Null length: error 94
        Case 13: v = Mid(12345, 2, 2)                 ' number as string argument
        Case 14: v = Mid$("abcdef", 2, 3)             ' $ form: String
        Case 15: v = Mid$(Null, 1)                    ' $ form with Null: error 94
        Case 16: v = Mid("abc", 2, Empty)             ' Empty length
        Case 17: v = Mid("abc", 1.5, 1)               ' fractional start rounds (banker's)
        Case 18: v = Mid("abc", "2", "1")             ' numeric strings for numeric args
        Case 19: v = Mid("", 1, 1)                    ' empty string
        Case 20: v = Mid(Empty, 1)                    ' Empty as the string
        Case 21: v = Mid(String:="abcdef", Start:=3, Length:=2)  ' named arguments
        Case 22: v = Mid$(String:="abcdef", Start:=3)            ' named arguments, $ form
        Case 23: v = Mid("abc", 2147483647)           ' largest Long start
    End Select
    T = "Mid." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Mid." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
