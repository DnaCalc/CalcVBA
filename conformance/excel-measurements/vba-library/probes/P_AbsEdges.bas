Attribute VB_Name = "P_AbsEdges"
Option Explicit
' Excel-backed Abs integer minimums, typed operands and Variant operands.
Private Const N As Long = 24
Private calls As Long
Private Function NextMinimum() As Integer
    calls = calls + 1
    NextMinimum = -32768
End Function
Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, i As Integer, l As Long, ll As LongLong, en As Long
    On Error GoTo fail
    Select Case k
        Case 1: i = -32768: v = Abs(i)
        Case 2: i = -32768: x = i: v = Abs(x)
        Case 3: v = Abs(CInt(-32768))
        Case 4: i = -32768: v = Abs(CVar(i))
        Case 5: v = Abs(-32768)
        Case 6: l = -2147483647 - 1: v = Abs(l)
        Case 7: l = -2147483647 - 1: x = l: v = Abs(x)
        Case 8: v = Abs(CLng(-2147483648#))
        Case 9: l = -2147483647 - 1: v = Abs(CVar(l))
        Case 10: x = CLngLng("-9223372036854775808"): v = Abs(x)
        Case 11: ll = CLngLng("-9223372036854775808"): v = Abs(ll)
        Case 12: v = Abs(CLngLng("-9223372036854775808"))
        Case 13: x = -2147483648#: v = Abs(x)
        Case 14: i = -32767: v = Abs(i)
        Case 15: l = -2147483647: v = Abs(l)
        Case 16: v = Abs(CDec(-32768))
        Case 17: v = Abs(CDate(-32768))
        Case 18: v = Abs(CInt(32767))
        Case 19: ll = CLngLng("-9223372036854775807"): v = Abs(ll)
        Case 20: v = Abs(CVErr(5))
        Case 21
            calls = 0
            On Error Resume Next
            v = Abs(NextMinimum())
            en = Err.Number
            Err.Clear
            On Error GoTo fail
            v = CStr(en) & ":" & CStr(calls)
        Case 22: v = VBA.Abs(CInt(-32768))
        Case 23: v = VBA.Math.Abs(CVar(CInt(-32768)))
        Case 24: v = Abs((CInt(-32768)))
    End Select
    T = "AbsEdges." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "AbsEdges." & k & "=ERR" & Err.Number
End Function
Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
