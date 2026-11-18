Attribute VB_Name = "P_FinancialEdges"
Option Explicit
' UNMEASURED cash-flow shape, leading zeros, convergence and rate boundaries.
Private Const N As Long = 22
Private Function T(ByVal k As Long) As String
    Dim v As Variant, d() As Double, m(0 To 1, 0 To 1) As Double
    On Error GoTo fail
    ReDim d(-1 To 1)
    d(-1) = 0: d(0) = -100: d(1) = 110
    m(0, 0) = -100: m(1, 0) = 30: m(0, 1) = 40: m(1, 1) = 50
    Select Case k
        Case 1: v = IRR(d)
        Case 2: d(-1) = -100: d(0) = 100: d(1) = 0: v = IRR(d)
        Case 3: d(-1) = -100: d(0) = 0: d(1) = 100: v = IRR(d)
        Case 4: v = IRR(d, -0.99999)
        Case 5: v = IRR(d, -0.999999)
        Case 6: v = IRR(d, -2)
        Case 7: v = IRR(d, 1E+300)
        Case 8: d(-1) = -1E+300: d(0) = 1E+300: d(1) = 1E+300: v = IRR(d)
        Case 9: d(-1) = -1E-300: d(0) = 1E-300: d(1) = 1E-300: v = IRR(d)
        Case 10: v = IRR(m)
        Case 11: v = MIRR(m, 0.1, 0.1)
        Case 12: v = MIRR(d, -2, 0.1)
        Case 13: v = MIRR(d, 0.1, -2)
        Case 14: v = MIRR(d, Null, 0.1)
        Case 15: v = MIRR(d, 0.1, Null)
        Case 16: v = MIRR(d, Empty, Empty)
        Case 17: v = MIRR(d, 1E+300, 1E+300)
        Case 18: v = MIRR(d, -1.00001, 0.1)
        Case 19: ReDim d(0 To 0): d(0) = 0: v = MIRR(d, -1, 0.1)
        Case 20: Erase d: v = MIRR(d, Null, 0.1)
        Case 21: Erase d: v = IRR(d, Null)
        Case 22: d(-1) = 0: d(0) = 0: d(1) = 0: v = IRR(d, 0)
    End Select
    T = "FinancialEdges." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "FinancialEdges." & k & "=ERR" & Err.Number
End Function
Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
