Attribute VB_Name = "P_WeekdayName"
Option Explicit
' WeekdayName (Strings; names are English on this machine). Weekday is relative to FirstDayOfWeek.
Private Const N As Long = 21

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = WeekdayName(1)                    ' default first day. locale
        Case 2: v = WeekdayName(7)                    ' locale
        Case 3: v = WeekdayName(1, True)              ' abbreviated. locale
        Case 4: v = WeekdayName(1, False, vbSunday)   ' Sunday first
        Case 5: v = WeekdayName(1, False, vbMonday)   ' Monday first
        Case 6: v = WeekdayName(7, False, vbMonday)   ' last day when Monday is first
        Case 7: v = WeekdayName(1, True, vbSaturday)  ' Saturday first, abbreviated
        Case 8: v = WeekdayName(1, False, vbUseSystemDayOfWeek)  ' system first day. locale
        Case 9: v = WeekdayName(0)                    ' error 5
        Case 10: v = WeekdayName(8)                   ' error 5
        Case 11: v = WeekdayName(-1)                  ' error 5
        Case 12: v = WeekdayName(1, False, 8)         ' invalid first day: error 5
        Case 13: v = WeekdayName(1, False, -1)        ' invalid first day: error 5
        Case 14: v = WeekdayName(Null)                ' error 94
        Case 15: v = WeekdayName(1, Null)             ' Null abbreviate: error 94
        Case 16: v = WeekdayName("2", False, vbSunday)  ' numeric string
        Case 17: v = WeekdayName(2.5, False, vbSunday)  ' fractional rounds (banker's)
        Case 18: v = WeekdayName(Weekday:=3, Abbreviate:=True, FirstDayOfWeek:=vbSunday)  ' named arguments
        Case 19: v = WeekdayName(1, 1, vbSunday)      ' numeric Abbreviate
        Case 20: v = WeekdayName(1, False, Null)      ' Null first day: error 94
        Case 21: v = TypeName(WeekdayName(1, False, vbSunday))  ' result type
    End Select
    T = "WeekdayName." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "WeekdayName." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
