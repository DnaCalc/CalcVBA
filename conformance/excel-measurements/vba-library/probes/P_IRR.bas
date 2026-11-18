Attribute VB_Name = "P_IRR"
Option Explicit
' IRR (VBA.Financial).
Private Const N As Long = 22

Private Function MK(ByVal lb As Long, ParamArray a() As Variant) As Double()
    Dim r() As Double, i As Long
    ReDim r(lb To lb + UBound(a))
    For i = 0 To UBound(a)
        r(lb + i) = a(i)
    Next
    MK = r
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    Dim d() As Double
    On Error GoTo fail
    Select Case k
        Case 1: d = MK(0, -100, 30, 40, 50, 60): v = IRR(d)           ' typical
        Case 2: d = MK(0, -100, 30, 40, 50, 60): v = IRR(d, 0.5)      ' guess 0.5
        Case 3: d = MK(0, -100, 30, 40, 50, 60): v = IRR(d, -0.5)     ' negative guess
        Case 4: d = MK(0, -100, 110): v = IRR(d)                      ' two values: 0.1
        Case 5: d = MK(0, -100, 0, 0, 121): v = IRR(d)                ' sparse
        Case 6: d = MK(0, 10, 20, 30): v = IRR(d)                     ' all positive: error 5
        Case 7: d = MK(0, -10, -20, -30): v = IRR(d)                  ' all negative: error 5
        Case 8: d = MK(0, -100): v = IRR(d)                           ' single value: error 5
        Case 9: d = MK(0, 0, 0, 0): v = IRR(d)                        ' all zero: error 5
        Case 10: d = MK(0, -1000, 1): v = IRR(d)                      ' return near -100%
        Case 11: d = MK(0, -100, 300, -300, 150): v = IRR(d)          ' multiple sign changes
        Case 12: d = MK(1, -100, 10, 10, 10, 10, 110): v = IRR(d)     ' lower bound 1
        Case 13: d = MK(0, -100, 30, 40, 50, 60): v = IRR(d, 100)     ' huge guess
        Case 14: d = MK(0, -100, 30, 40, 50, 60): v = IRR(d, -1)      ' guess -1
        Case 15: d = MK(0, -100, 30, 40, 50, 60): v = IRR(d, 0)       ' guess 0
        Case 16: d = MK(0, -100, 1000): v = IRR(d)                    ' very large return
        Case 17: d = MK(0, -123.45, 33.1, 44.2, 55.3, 66.4): v = IRR(d)' awkward decimals
        Case 18: d = MK(0, -100, 30, 40, 50, 60): v = IRR(d, Null)    ' Null guess
        Case 19: d = MK(0, -100, 30, 40, 50, 60): v = IRR(ValueArray:=d, Guess:=0.2)' named
        Case 20: d = MK(0, 100, -50, -40, -30): v = IRR(d)            ' positive first then negative
        Case 21: d = MK(0, -100, 0.0001): v = IRR(d)                  ' return near -100%, tiny
        Case 22: v = IRR(d)                                           ' unallocated array
    End Select
    T = "IRR." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "IRR." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
