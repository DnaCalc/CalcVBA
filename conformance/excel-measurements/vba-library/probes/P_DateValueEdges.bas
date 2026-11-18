Attribute VB_Name = "P_DateValueEdges"
Option Explicit
' Excel-backed DateValue text coercion and rounded Date boundaries.
Private Const N As Long = 22
Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = DateValue(Null)
        Case 2: v = DateValue(Empty)
        Case 3: v = DateValue(True)
        Case 4: v = DateValue(False)
        Case 5: v = DateValue(CVErr(5))
        Case 6: v = DateValue(0)
        Case 7: v = DateValue(0.5)
        Case 8: v = DateValue(1.5)
        Case 9: v = DateValue(123)
        Case 10: v = DateValue("43831")
        Case 11: v = DateValue("43831.5")
        Case 12: v = DateValue("0.5")
        Case 13: v = DateValue("13:30 PM")
        Case 14: v = DateValue(CDate(-1.5))
        Case 15: v = DateValue(CDate(-0.25))
        Case 16: v = DateValue(CDate(43831.999999))
        Case 17: v = DateValue(CDate(0.99999999))
        Case 18: v = DateValue("1/1/10000")
        Case 19: v = DateValue("1/1/99")
        Case 20: v = DateValue("1/1/100")
        Case 21: v = DateValue("1/2/2020 10 AM")
        Case 22: v = DateValue("1/2/2020 1:02:03.5")
    End Select
    T = "DateValueEdges." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "DateValueEdges." & k & "=ERR" & Err.Number
End Function
Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
