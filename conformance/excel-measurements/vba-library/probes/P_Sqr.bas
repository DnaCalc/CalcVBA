Attribute VB_Name = "P_Sqr"
Option Explicit
' Sqr (VBA.Math).
Private Const N As Long = 23

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Sqr(0)                                            ' zero
        Case 2: v = Sqr(1)                                            ' one
        Case 3: v = Sqr(2)                                            ' two
        Case 4: v = Sqr(4)                                            ' four
        Case 5: v = Sqr(0.25)                                         ' quarter
        Case 6: v = Sqr(3)                                            ' three
        Case 7: v = Sqr(10)                                           ' ten
        Case 8: v = Sqr(1E-300)                                       ' tiny
        Case 9: v = Sqr(1E+300)                                       ' huge
        Case 10: v = Sqr(1.7976931348623157E+308)                     ' max double
        Case 11: v = Sqr(4.94065645841247E-324)                       ' smallest denormal
        Case 12: v = Sqr(1.0000000000000002)                          ' 1+eps
        Case 13: v = Sqr(123456789)                                   ' large integer
        Case 14: v = Sqr(15241578750190521#)                          ' perfect square 123456789^2
        Case 15: v = Sqr(-1)                                          ' negative: error 5
        Case 16: v = Sqr(-1E-300)                                     ' tiny negative: error 5
        Case 17: v = Sqr(0.1)                                         ' 0.1
        Case 18: v = Sqr("0.5")                                       ' numeric string
        Case 19: v = Sqr(Null)                                        ' Null: error 94
        Case 20: v = Sqr(Empty)                                       ' Empty is 0
        Case 21: v = Sqr(True)                                        ' Boolean True is -1
        Case 22: v = Sqr("abc")                                       ' non-numeric string: error 13
        Case 23: v = Sqr(Number:=1)                                   ' named argument
    End Select
    T = "Sqr." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Sqr." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
