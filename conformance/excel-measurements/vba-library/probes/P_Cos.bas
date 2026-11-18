Attribute VB_Name = "P_Cos"
Option Explicit
' Cos (VBA.Math).
Private Const N As Long = 24

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Cos(0)                                            ' zero
        Case 2: v = Cos(1)                                            ' one
        Case 3: v = Cos(-1)                                           ' minus one
        Case 4: v = Cos(0.5)                                          ' half
        Case 5: v = Cos(1E-300)                                       ' tiny
        Case 6: v = Cos(1E+300)                                       ' huge
        Case 7: v = Cos(1E+15)                                        ' large
        Case 8: v = Cos(10)                                           ' ten
        Case 9: v = Cos(3.141592653589793)                            ' pi
        Case 10: v = Cos(3.14159265358979)                            ' pi 15 digits
        Case 11: v = Cos(1.5707963267948966)                          ' pi/2
        Case 12: v = Cos(4.71238898038469)                            ' 3 pi/2
        Case 13: v = Cos(6.283185307179586)                           ' 2 pi
        Case 14: v = Cos(1.0471975511965976)                          ' pi/3
        Case 15: v = Cos(0.7853981633974483)                          ' pi/4
        Case 16: v = Cos(-1.5707963267948966)                         ' minus pi/2
        Case 17: v = Cos(1E+22)                                       ' 1e22
        Case 18: v = Cos(100)                                         ' hundred
        Case 19: v = Cos("0.5")                                       ' numeric string
        Case 20: v = Cos(Null)                                        ' Null: error 94
        Case 21: v = Cos(Empty)                                       ' Empty is 0
        Case 22: v = Cos(True)                                        ' Boolean True is -1
        Case 23: v = Cos("abc")                                       ' non-numeric string: error 13
        Case 24: v = Cos(Number:=1)                                   ' named argument
    End Select
    T = "Cos." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Cos." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
