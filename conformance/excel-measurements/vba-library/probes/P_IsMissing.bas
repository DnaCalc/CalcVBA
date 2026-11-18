Attribute VB_Name = "P_IsMissing"
Option Explicit
' IsMissing (Information; only meaningful for an Optional Variant parameter).
Private Const N As Long = 16

Private Function H1(Optional x As Variant) As Boolean
    H1 = IsMissing(x)
End Function
Private Function H2(Optional x As Long) As Boolean
    H2 = IsMissing(x)                     ' not a Variant: always False
End Function
Private Function H3(Optional x As Variant = 5) As Boolean
    H3 = IsMissing(x)                     ' has a default: never missing
End Function
Private Function H4(Optional x As Variant) As Boolean
    H4 = H1(x)                            ' Missing is passed along
End Function
Private Function H5(Optional ByVal x As Variant) As Boolean
    H5 = IsMissing(x)
End Function
Private Function H6(Optional x As Variant) As Boolean
    If IsMissing(x) Then x = 5
    H6 = IsMissing(x)                     ' assigned inside: no longer missing
End Function
Private Function H7(ParamArray p() As Variant) As Boolean
    H7 = IsMissing(p)
End Function
Private Function H8(Optional x As String) As Boolean
    H8 = IsMissing(x)
End Function
Private Function H9(Optional x As Variant) As String
    H9 = TypeName(x) & "|" & VarType(x) & "|" & IsError(x) & "|" & IsEmpty(x)
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    Dim va As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = H1()                      ' not passed
        Case 2: v = H1(1)                     ' passed
        Case 3: v = H1(Empty)                 ' Empty passed explicitly
        Case 4: v = H1(Null)
        Case 5: v = H1("")
        Case 6: v = H2()                      ' Optional Long
        Case 7: v = H3()                      ' Optional with default
        Case 8: v = H4()                      ' forwarded Missing
        Case 9: v = H5()                      ' Optional ByVal Variant
        Case 10: v = H6()                      ' assigned inside the procedure
        Case 11: v = H7()                      ' ParamArray
        Case 12: v = H8()                      ' Optional String
        Case 13: v = H9()                      ' type of a Missing value
        Case 14: v = H1(CVErr(448))            ' an ordinary error value is not Missing
        Case 15: v = H1(x:=Empty)              ' named argument
        Case 16: v = IsMissing(va)             ' ordinary Variant variable
    End Select
    T = "IsMissing." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "IsMissing." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
