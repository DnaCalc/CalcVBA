Attribute VB_Name = "P_NPVEdges"
Option Explicit
' NPV (VBA.Financial).
Private Const N As Long = 29

Private Function MK(ByVal lb As Long, ParamArray a() As Variant) As Double()
    Dim r() As Double, i As Long
    ReDim r(lb To lb + UBound(a))
    For i = 0 To UBound(a)
        r(lb + i) = a(i)
    Next
    MK = r
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant, d() As Double
    On Error GoTo fail
    Select Case k
        Case 1: d = MK(-3, -100, 50, 60): v = NPV(0, d)
        Case 2: d = MK(-3, -100, 50, 60): v = NPV(0.1, d)
        Case 3: d = MK(-3, -100, 50, 60): v = NPV(-0.1, d)
        Case 4: d = MK(-3, -100, 50, 60): v = NPV(-0.5, d)
        Case 5: d = MK(-3, -100, 50, 60): v = NPV(-1, d)
        Case 6: d = MK(-3, -100, 50, 60): v = NPV(-2, d)
        Case 7: d = MK(-3, -100, 50, 60): v = NPV(-3, d)
        Case 8: d = MK(-3, -100, 50, 60): v = NPV(1E-12, d)
        Case 9: d = MK(-3, -100, 50, 60): v = NPV(-1E-12, d)
        Case 10: d = MK(-3, -100, 50, 60): v = NPV(0.0001, d)
        Case 11: d = MK(-3, -100, 50, 60): v = NPV(1000, d)
        Case 12: d = MK(-3, -100, 50, 60): v = NPV(1E100, d)
        Case 13: d = MK(-3, -100, 50, 60): v = NPV(1E-100, d)
        Case 14: d = MK(1, 1, 0, 0, 1): v = NPV(0.1, d)
        Case 15: d = MK(1, 10, -20, 30, -40): v = NPV(0.05, d)
        Case 16: d = MK(1, 1.1, 2.2, 3.3, 4.4): v = NPV(0.07, d)
        Case 17: d = MK(1, 1E300, 1E300): v = NPV(0.5, d)
        Case 18: d = MK(1, 1E16, -1E16, 1): v = NPV(1E-12, d)
        Case 19: d = MK(1, 0, 0, 0): v = NPV(0.1, d)
        Case 20: d = MK(1, 100): v = NPV(0.1, d)
        Case 21: d = MK(1, -100): v = NPV(0.1, d)
        Case 22: d = MK(1, 1, 2, 3): v = NPV(-0.999999999999, d)
        Case 23: d = MK(1, 1E300, 1E300): v = NPV(1E300, d)
        Case 24: v = NPV(0.1, d)
        Case 25: v = NPV(-1, d)
        Case 26: v = NPV(0, d)
        Case 27: d = MK(0, 1, 2): v = NPV(Null, d)
        Case 28: d = MK(0, 1, 2): v = NPV(Empty, d)
        Case 29: d = MK(0, 1, 2): v = NPV("0.1", d)
    End Select
    T = "NPVEdges." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "NPVEdges." & k & "=ERR" & Err.Number
End Function
Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
