Attribute VB_Name = "P_MIRR"
Option Explicit
' MIRR (VBA.Financial).
Private Const N As Long = 18

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
        Case 1: d = MK(0, -120000, 39000, 30000, 21000, 37000, 46000): v = MIRR(d, 0.1, 0.12)' typical
        Case 2: d = MK(0, -120000, 39000, 30000, 21000, 37000, 46000): v = MIRR(d, 0.1, 0)' reinvest 0
        Case 3: d = MK(0, -120000, 39000, 30000, 21000, 37000, 46000): v = MIRR(d, 0, 0.12)' finance 0
        Case 4: d = MK(0, -120000, 39000, 30000, 21000, 37000, 46000): v = MIRR(d, -1, 0.12)' finance -1: error 5
        Case 5: d = MK(0, -120000, 39000, 30000, 21000, 37000, 46000): v = MIRR(d, 0.1, -1)' reinvest -1: error 5
        Case 6: d = MK(0, -120000, 39000, 30000, 21000, 37000, 46000): v = MIRR(d, -0.5, -0.5)' negative rates
        Case 7: d = MK(0, 10, 20, 30): v = MIRR(d, 0.1, 0.12)         ' all positive: error 5
        Case 8: d = MK(0, -10, -20, -30): v = MIRR(d, 0.1, 0.12)      ' all negative: error 5
        Case 9: d = MK(0, -100): v = MIRR(d, 0.1, 0.12)               ' single value: error 5
        Case 10: d = MK(0, -100, 110): v = MIRR(d, 0.1, 0.1)          ' two values
        Case 11: d = MK(1, -100, 30, 40, 50, 60): v = MIRR(d, 0.08, 0.09)' lower bound 1
        Case 12: d = MK(0, -100, 300, -300, 150): v = MIRR(d, 0.1, 0.12)' multiple sign changes
        Case 13: d = MK(0, -123.45, 33.1, 44.2, 55.3, 66.4): v = MIRR(d, 0.0725, 0.0825)' awkward decimals
        Case 14: d = MK(0, -1000, 1): v = MIRR(d, 0.1, 0.1)           ' strong loss
        Case 15: d = MK(0, -100, 50, 60): v = MIRR(d, 1000, 1000)     ' huge rates
        Case 16: d = MK(0, 0, 0, 0): v = MIRR(d, 0.1, 0.1)            ' all zero
        Case 17: v = MIRR(d, 0.1, 0.1)                                ' unallocated array
        Case 18: d = MK(0, -100, 50, 60): v = MIRR(ValueArray:=d, FinanceRate:=0.1, ReinvestRate:=0.12)' named
    End Select
    T = "MIRR." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "MIRR." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
