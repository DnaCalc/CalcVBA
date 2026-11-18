Attribute VB_Name = "P_DateConversionEdges"
Option Explicit
' Excel-backed CDate/CVDate Variant and conversion boundaries.
Private Const N As Long = 47
Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = CDate(Null)
        Case 2: v = CDate(Empty)
        Case 3: v = CDate(CVErr(5))
        Case 4: v = CDate(CVErr(32767))
        Case 5: v = CDate(2958466)
        Case 6: v = CDate(-657435)
        Case 7: v = CDate(2958465.999999999)
        Case 8: v = CDate("43831")
        Case 9: v = CDate("43831.5")
        Case 10: v = CDate(" 43831.5 ")
        Case 11: v = CDate("2958466")
        Case 12: v = CDate("1e3")
        Case 13: v = CDate("&H10")
        Case 14: v = CDate("+1.5")
        Case 15: v = CDate("-1.5")
        Case 16: v = CDate("0.5")
        Case 17: v = CDate("1")
        Case 18: v = CDate("abc")
        Case 19: v = CDate("12:30")
        Case 20: v = CDate(True)
        Case 21: v = CVDate(Null)
        Case 22: v = CVDate(Empty)
        Case 23: v = CVDate(CVErr(5))
        Case 24: v = CVDate(CVErr(32767))
        Case 25: v = CVDate(2958466)
        Case 26: v = CVDate(-657435)
        Case 27: v = CVDate(2958465.999999999)
        Case 28: v = CVDate("43831")
        Case 29: v = CVDate("43831.5")
        Case 30: v = CVDate(" 43831.5 ")
        Case 31: v = CVDate("2958466")
        Case 32: v = CVDate("1e3")
        Case 33: v = CVDate("&H10")
        Case 34: v = CVDate("+1.5")
        Case 35: v = CVDate("-1.5")
        Case 36: v = CVDate("0.5")
        Case 37: v = CVDate("1")
        Case 38: v = CVDate("abc")
        Case 39: v = CVDate("12:30")
        Case 40: v = CVDate(True)
        Case 41: v = VBA.CVDate(Null)
        Case 42: v = VBA.Conversion.CVDate(Null)
        Case 43: v = Len(CVDate(1))
        Case 44: v = CDate("12/29/1899 12:00 PM")
        Case 45: v = CVDate("12/29/1899 12:00 PM")
        Case 46: v = CDate("1/1/100 6:00 AM")
        Case 47: v = CVDate("12/30/1899 12:00 PM")
    End Select
    T = "DateConversionEdges." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "DateConversionEdges." & k & "=ERR" & Err.Number
End Function
Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
