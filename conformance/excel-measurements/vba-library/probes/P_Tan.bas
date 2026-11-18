Attribute VB_Name = "P_Tan"
Option Explicit
' Tan (VBA.Math).
Private Const N As Long = 22

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Tan(0)                                            ' zero
        Case 2: v = Tan(1)                                            ' one
        Case 3: v = Tan(-1)                                           ' minus one
        Case 4: v = Tan(0.5)                                          ' half
        Case 5: v = Tan(1E-300)                                       ' tiny
        Case 6: v = Tan(1E+300)                                       ' huge
        Case 7: v = Tan(1E+15)                                        ' large
        Case 8: v = Tan(1.5707963267948966)                           ' pi/2: huge, not infinite
        Case 9: v = Tan(-1.5707963267948966)                          ' minus pi/2
        Case 10: v = Tan(0.7853981633974483)                          ' pi/4: near 1
        Case 11: v = Tan(3.141592653589793)                           ' pi
        Case 12: v = Tan(6.283185307179586)                           ' 2 pi
        Case 13: v = Tan(4.71238898038469)                            ' 3 pi/2
        Case 14: v = Tan(100)                                         ' hundred
        Case 15: v = Tan(1.5707963)                                   ' pi/2 truncated
        Case 16: v = Tan(1E+22)                                       ' 1e22
        Case 17: v = Tan("0.5")                                       ' numeric string
        Case 18: v = Tan(Null)                                        ' Null: error 94
        Case 19: v = Tan(Empty)                                       ' Empty is 0
        Case 20: v = Tan(True)                                        ' Boolean True is -1
        Case 21: v = Tan("abc")                                       ' non-numeric string: error 13
        Case 22: v = Tan(Number:=1)                                   ' named argument
    End Select
    T = "Tan." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Tan." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
