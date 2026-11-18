Attribute VB_Name = "P_Atn"
Option Explicit
' Atn (VBA.Math).
Private Const N As Long = 21

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Atn(0)                                            ' zero
        Case 2: v = Atn(1)                                            ' one: pi/4
        Case 3: v = Atn(-1)                                           ' minus one
        Case 4: v = Atn(0.5)                                          ' half
        Case 5: v = Atn(2)                                            ' two
        Case 6: v = Atn(10)                                           ' ten
        Case 7: v = Atn(1E-300)                                       ' tiny
        Case 8: v = Atn(1E+300)                                       ' huge: pi/2
        Case 9: v = Atn(-1E+300)                                      ' huge negative
        Case 10: v = Atn(1E+15)                                       ' large
        Case 11: v = Atn(0.1)                                         ' 0.1
        Case 12: v = Atn(1.0000000000000002)                          ' just above 1
        Case 13: v = Atn(0.9999999999999999)                          ' just below 1
        Case 14: v = Atn(0.7071067811865476)                          ' 1/sqrt2
        Case 15: v = Atn(1.7320508075688772)                          ' sqrt3: pi/3
        Case 16: v = Atn("0.5")                                       ' numeric string
        Case 17: v = Atn(Null)                                        ' Null: error 94
        Case 18: v = Atn(Empty)                                       ' Empty is 0
        Case 19: v = Atn(True)                                        ' Boolean True is -1
        Case 20: v = Atn("abc")                                       ' non-numeric string: error 13
        Case 21: v = Atn(Number:=1)                                   ' named argument
    End Select
    T = "Atn." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Atn." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
