Attribute VB_Name = "P_IsError"
Option Explicit
' IsError (Information).
Private Const N As Long = 16

Private Function HM(Optional x As Variant) As Boolean
    HM = IsError(x)
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    Dim va As Variant, z As Long
    On Error GoTo fail
    Select Case k
        Case 1: v = IsError(CVErr(5))                 ' CVErr value
        Case 2: v = IsError(CVErr(0))                 ' error number 0
        Case 3: v = IsError(CVErr(2042))              ' #N/A as an error value
        Case 4: va = CVErr(11): v = IsError(va)       ' Variant holding an error
        Case 5: v = IsError("#N/A")                   ' string that looks like an error
        Case 6: v = IsError(1)
        Case 7: v = IsError(Null)
        Case 8: v = IsError(Empty)
        Case 9: v = IsError(Array(CVErr(1)))          ' array of errors is not an error
        Case 10: v = IsError(Array(CVErr(1))(0))       ' element that is an error
        Case 11: v = HM()                              ' Missing argument is an error value
        Case 12: v = HM(Empty)
        Case 13: v = IsError(CVErr(65535))             ' largest CVErr number
        Case 14: v = IsError(1 \ z)                    ' failing expression raises before IsError
        Case 15: v = IsError(CVErr(5) = CVErr(5))      ' comparison of errors: error 13
        Case 16: va = CVErr(7): v = IsError(va) & "|" & CLng(va) & "|" & VarType(va)   ' number round trip
    End Select
    T = "IsError." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "IsError." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
