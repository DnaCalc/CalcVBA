Attribute VB_Name = "P_Exp"
Option Explicit
' Exp (VBA.Math).
Private Const N As Long = 27

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Exp(0)                                            ' zero
        Case 2: v = Exp(1)                                            ' one: e
        Case 3: v = Exp(-1)                                           ' minus one
        Case 4: v = Exp(0.5)                                          ' half
        Case 5: v = Exp(1E-300)                                       ' tiny
        Case 6: v = Exp(-1E-300)                                      ' tiny negative
        Case 7: v = Exp(2)                                            ' two
        Case 8: v = Exp(10)                                           ' ten
        Case 9: v = Exp(100)                                          ' hundred
        Case 10: v = Exp(-100)                                        ' minus hundred
        Case 11: v = Exp(709)                                         ' just below overflow
        Case 12: v = Exp(709.782712893384)                            ' largest finite
        Case 13: v = Exp(709.7827128933841)                           ' at the overflow edge
        Case 14: v = Exp(710)                                         ' overflow: error 6
        Case 15: v = Exp(1E+300)                                      ' huge: error 6
        Case 16: v = Exp(-745)                                        ' near underflow
        Case 17: v = Exp(-746)                                        ' underflow to 0
        Case 18: v = Exp(-1E+300)                                     ' huge negative: 0
        Case 19: v = Exp(2.302585092994046)                           ' ln 10
        Case 20: v = Exp(36.04365338911715)                           ' ~1e15
        Case 21: v = Exp(0.6931471805599453)                          ' ln 2
        Case 22: v = Exp("0.5")                                       ' numeric string
        Case 23: v = Exp(Null)                                        ' Null: error 94
        Case 24: v = Exp(Empty)                                       ' Empty is 0
        Case 25: v = Exp(True)                                        ' Boolean True is -1
        Case 26: v = Exp("abc")                                       ' non-numeric string: error 13
        Case 27: v = Exp(Number:=1)                                   ' named argument
    End Select
    T = "Exp." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Exp." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
