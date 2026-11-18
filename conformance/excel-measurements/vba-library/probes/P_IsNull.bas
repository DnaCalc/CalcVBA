Attribute VB_Name = "P_IsNull"
Option Explicit
' IsNull (Information).
Private Const N As Long = 20

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    Dim va As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = IsNull(Null)                      ' Null literal
        Case 2: v = IsNull(Empty)                     ' Empty
        Case 3: v = IsNull("")                        ' zero-length string
        Case 4: v = IsNull(0)                         ' zero
        Case 5: v = IsNull(Null + 1)                  ' arithmetic propagates Null
        Case 6: v = IsNull(Null & "a")                ' concatenation does not
        Case 7: v = IsNull(Null And False)            ' Null And False = False
        Case 8: v = IsNull(Null And True)             ' Null And True = Null
        Case 9: v = IsNull(Null Or True)              ' Null Or True = True
        Case 10: v = IsNull(Null Or False)             ' Null Or False = Null
        Case 11: v = IsNull(Left(Null, 1))             ' Left propagates Null
        Case 12: v = IsNull(UCase(Null))               ' UCase propagates Null
        Case 13: v = IsNull(Null = 1)                  ' comparison with Null
        Case 14: v = IsNull(IIf(True, Null, 1))        ' IIf returning Null
        Case 15: va = Null: v = IsNull(va)             ' Variant holding Null
        Case 16: Set va = Nothing: v = IsNull(va)      ' Nothing
        Case 17: v = IsNull(Array(Null))               ' array containing Null
        Case 18: v = IsNull(Array(Null)(0))            ' element that is Null
        Case 19: v = IsNull(CStr(Null))                ' CStr(Null): error 94
        Case 20: s = "x": v = IsNull(s)                ' String variable
    End Select
    T = "IsNull." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "IsNull." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
