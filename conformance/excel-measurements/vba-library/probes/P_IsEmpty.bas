Attribute VB_Name = "P_IsEmpty"
Option Explicit
' IsEmpty (Information).
Private Const N As Long = 16

Private Function HE(Optional x As Variant) As Boolean
    HE = IsEmpty(x)
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    Dim va As Variant, s As String, l As Long, o As Object, fa(1 To 2) As Long, d As Date
    On Error GoTo fail
    Select Case k
        Case 1: v = IsEmpty(va)                       ' uninitialised Variant
        Case 2: v = IsEmpty(Empty)                    ' Empty literal
        Case 3: va = 0: v = IsEmpty(va)               ' zero
        Case 4: va = "": v = IsEmpty(va)              ' zero-length string
        Case 5: va = Null: v = IsEmpty(va)            ' Null
        Case 6: Set va = Nothing: v = IsEmpty(va)     ' Nothing in a Variant
        Case 7: v = IsEmpty(vbNullString)             ' vbNullString
        Case 8: v = IsEmpty(s)                        ' String variable
        Case 9: v = IsEmpty(l)                        ' Long variable
        Case 10: v = IsEmpty(o)                        ' unset Object variable
        Case 11: v = IsEmpty(fa)                       ' array
        Case 12: va = 5: va = Empty: v = IsEmpty(va)   ' assigned Empty after a value
        Case 13: va = CVErr(5): v = IsEmpty(va)        ' error value
        Case 14: v = HE()                              ' Missing argument is not Empty
        Case 15: v = HE(Empty)                         ' Empty passed explicitly
        Case 16: v = IsEmpty(d)                        ' Date variable (0)
    End Select
    T = "IsEmpty." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "IsEmpty." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
