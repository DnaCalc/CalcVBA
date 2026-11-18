Attribute VB_Name = "P_NPV"
Option Explicit
' NPV (VBA.Financial).
Private Const N As Long = 16

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
        Case 1: d = MK(0, -1000, 300, 400, 500, 600): v = NPV(0.1, d) ' typical
        Case 2: d = MK(0, -1000, 300, 400, 500, 600): v = NPV(0, d)   ' zero rate
        Case 3: d = MK(0, 100): v = NPV(0.1, d)                       ' single value
        Case 4: d = MK(1, -1000, 300, 400, 500, 600): v = NPV(0.1, d) ' array lower bound 1
        Case 5: d = MK(0, 10, 20, 30): v = NPV(0.05, d)               ' all positive
        Case 6: d = MK(0, -10, -20, -30): v = NPV(0.05, d)            ' all negative
        Case 7: d = MK(0, -1000, 300, 400): v = NPV(-0.5, d)          ' negative rate
        Case 8: d = MK(0, -1000, 300, 400): v = NPV(-1, d)            ' rate -1: error 5
        Case 9: d = MK(0, -1000, 300, 400): v = NPV(-2, d)            ' rate below -1
        Case 10: d = MK(0, 1.1, 2.2, 3.3, 4.4): v = NPV(0.07, d)      ' awkward decimals
        Case 11: d = MK(0, 1E+300, 1E+300): v = NPV(0.5, d)           ' huge values
        Case 12: d = MK(0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1): v = NPV(0.0001, d)' twelve values, small rate
        Case 13: v = NPV(0.1, d)                                      ' unallocated array
        Case 14: d = MK(0, -1000, 300, 400, 500, 600): v = NPV(Rate:=0.1, ValueArray:=d)' named
        Case 15: d = MK(0, -100, 50, 60): v = NPV(1E-12, d)           ' tiny rate
        Case 16: d = MK(0, -100, 50, 60): v = NPV(1000, d)            ' huge rate
    End Select
    T = "NPV." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "NPV." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
