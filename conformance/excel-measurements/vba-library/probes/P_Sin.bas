Attribute VB_Name = "P_Sin"
Option Explicit
' Sin (VBA.Math).
Private Const N As Long = 25

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Sin(0)                                            ' zero
        Case 2: v = Sin(1)                                            ' one
        Case 3: v = Sin(-1)                                           ' minus one
        Case 4: v = Sin(0.5)                                          ' half
        Case 5: v = Sin(1E-300)                                       ' tiny
        Case 6: v = Sin(-1E-300)                                      ' tiny negative
        Case 7: v = Sin(1E+300)                                       ' huge
        Case 8: v = Sin(1E+15)                                        ' large
        Case 9: v = Sin(10)                                           ' ten
        Case 10: v = Sin(100)                                         ' hundred
        Case 11: v = Sin(3.141592653589793)                           ' pi
        Case 12: v = Sin(3.14159265358979)                            ' pi 15 digits
        Case 13: v = Sin(1.5707963267948966)                          ' pi/2
        Case 14: v = Sin(6.283185307179586)                           ' 2 pi
        Case 15: v = Sin(31.41592653589793)                           ' 10 pi
        Case 16: v = Sin((355# / 113#))                               ' near pi
        Case 17: v = Sin(0.1)                                         ' 0.1
        Case 18: v = Sin(-3.141592653589793)                          ' minus pi
        Case 19: v = Sin(1E+22)                                       ' 1e22
        Case 20: v = Sin("0.5")                                       ' numeric string
        Case 21: v = Sin(Null)                                        ' Null: error 94
        Case 22: v = Sin(Empty)                                       ' Empty is 0
        Case 23: v = Sin(True)                                        ' Boolean True is -1
        Case 24: v = Sin("abc")                                       ' non-numeric string: error 13
        Case 25: v = Sin(Number:=1)                                   ' named argument
    End Select
    T = "Sin." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Sin." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
