Attribute VB_Name = "P_TimeValueEdges"
Option Explicit
' Excel-backed TimeValue coercion, second rounding and meridian boundaries.
Private Const N As Long = 33
Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = TimeValue(0)
        Case 2: v = TimeValue(0.5)
        Case 3: v = TimeValue(1)
        Case 4: v = TimeValue(1.5)
        Case 5: v = TimeValue(True)
        Case 6: v = TimeValue(CVErr(5))
        Case 7: v = TimeValue(CDate(-1.5))
        Case 8: v = TimeValue(CDate(43831.123456))
        Case 9: v = TimeValue(CDate(0.99999999))
        Case 10: v = TimeValue("0:30 AM")
        Case 11: v = TimeValue("0:30 PM")
        Case 12: v = TimeValue("13:30 AM")
        Case 13: v = TimeValue("13:30 PM")
        Case 14: v = TimeValue("23:30 PM")
        Case 15: v = TimeValue("24:30 PM")
        Case 16: v = TimeValue("12:30 AM")
        Case 17: v = TimeValue("1:02:03.0")
        Case 18: v = TimeValue("1:02:03.5")
        Case 19: v = TimeValue(" 1 : 2 : 3 ")
        Case 20: v = TimeValue("1/2/2020 10 AM")
        Case 21: v = TimeValue("1/2/2020 13:30 PM")
        Case 22: v = TimeValue("0.5")
        Case 23: v = TimeValue("43831.5")
        Case 24: v = TimeValue(CDate(-0.25))
        Case 25: v = TimeValue(CDate(0.000005))
        Case 26: v = TimeValue(CDate(0.000006))
        Case 27: v = TimeValue(CDate(43831.999999))
        Case 28: v = TimeValue("23:59:59 PM")
        Case 29: v = TimeValue("23:59:59 AM")
        Case 30: v = TimeValue("-1:00")
        Case 31: v = TimeValue(Empty)
        Case 32: v = TimeValue(Null)
        Case 33: v = TimeValue(#2020-01-16# + 0.5515625)
    End Select
    T = "TimeValueEdges." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "TimeValueEdges." & k & "=ERR" & Err.Number
End Function
Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
