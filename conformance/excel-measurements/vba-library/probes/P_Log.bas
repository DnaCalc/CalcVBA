Attribute VB_Name = "P_Log"
Option Explicit
' Log (VBA.Math).
Private Const N As Long = 24

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Log(1)                                            ' one: 0
        Case 2: v = Log(2.718281828459045)                            ' e
        Case 3: v = Log(10)                                           ' ten
        Case 4: v = Log(0.5)                                          ' half
        Case 5: v = Log(2)                                            ' two
        Case 6: v = Log(100)                                          ' hundred
        Case 7: v = Log(1E-300)                                       ' tiny
        Case 8: v = Log(1E+300)                                       ' huge
        Case 9: v = Log(1E+308)                                       ' near max
        Case 10: v = Log(4.94065645841247E-324)                       ' smallest denormal
        Case 11: v = Log(1.0000000000000002)                          ' 1+eps
        Case 12: v = Log(0.9999999999999999)                          ' 1-eps/2
        Case 13: v = Log(1.0001)                                      ' just above 1
        Case 14: v = Log(0)                                           ' zero: error 5
        Case 15: v = Log(-1)                                          ' negative: error 5
        Case 16: v = Log(1E+15)                                       ' large
        Case 17: v = Log(7.38905609893065)                            ' e squared
        Case 18: v = Log(1000)                                        ' thousand
        Case 19: v = Log("0.5")                                       ' numeric string
        Case 20: v = Log(Null)                                        ' Null: error 94
        Case 21: v = Log(Empty)                                       ' Empty is 0
        Case 22: v = Log(True)                                        ' Boolean True is -1
        Case 23: v = Log("abc")                                       ' non-numeric string: error 13
        Case 24: v = Log(Number:=1)                                   ' named argument
    End Select
    T = "Log." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Log." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
