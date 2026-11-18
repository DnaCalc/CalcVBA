Attribute VB_Name = "P_Weekday"
Option Explicit
' Weekday (VBA.DateTime; MS-VBAL 6.1.2.x).
Private Const N As Long = 34

Private Function T(ByVal k As Long) As String
    Dim v As Variant, a As Variant, b As Variant, c As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Weekday(DateSerial(2020, 2, 29))   ' Saturday, default Sunday-first
        Case 2: v = Weekday(DateSerial(2020, 3, 1))   ' Sunday
        Case 3: v = Weekday(DateSerial(2020, 3, 2))   ' Monday
        Case 4: v = Weekday(DateSerial(2020, 2, 29), vbSunday)   ' first day Sunday
        Case 5: v = Weekday(DateSerial(2020, 2, 29), vbMonday)   ' first day Monday
        Case 6: v = Weekday(DateSerial(2020, 2, 29), vbTuesday)   ' first day Tuesday
        Case 7: v = Weekday(DateSerial(2020, 2, 29), vbWednesday)   ' first day Wednesday
        Case 8: v = Weekday(DateSerial(2020, 2, 29), vbThursday)   ' first day Thursday
        Case 9: v = Weekday(DateSerial(2020, 2, 29), vbFriday)   ' first day Friday
        Case 10: v = Weekday(DateSerial(2020, 2, 29), vbSaturday)   ' first day Saturday
        Case 11: v = Weekday(DateSerial(2020, 2, 29), vbUseSystemDayOfWeek)   ' system first day ' locale
        Case 12: v = Weekday(DateSerial(2020, 2, 29), 8)   ' out of range: error 5
        Case 13: v = Weekday(DateSerial(2020, 2, 29), -1)   ' negative: error 5
        Case 14: v = Weekday(DateSerial(2020, 2, 29), "2")   ' numeric string first day
        Case 15: v = Weekday(DateSerial(2020, 2, 29), 1.5)   ' fractional first day rounds to even
        Case 16: v = Weekday(DateSerial(2020, 2, 29), Null)   ' Null first day: error 94
        Case 17: v = Weekday(Null)   ' Null propagates
        Case 18: v = Weekday(Null, vbMonday)   ' Null with first day
        Case 19: v = Weekday(Empty)   ' Empty is serial 0
        Case 20: v = Weekday(0)   ' serial 0 is Saturday
        Case 21: v = Weekday(-1)   ' serial -1 is Friday
        Case 22: v = Weekday(-0.5)   ' negative fraction
        Case 23: v = Weekday(1.99)   ' fraction ignored
        Case 24: v = Weekday("2020/02/29")   ' string date ' locale
        Case 25: v = Weekday("February 29, 2020")   ' string date with month name
        Case 26: v = Weekday("abc")   ' invalid string: error 13
        Case 27: v = Weekday("")   ' empty string: error 13
        Case 28: v = Weekday(True)   ' Boolean -1
        Case 29: v = Weekday(DateSerial(9999, 12, 31))   ' upper limit
        Case 30: v = Weekday(DateSerial(100, 1, 1))   ' lower limit
        Case 31: v = Weekday(2958466)   ' beyond upper limit: error 5
        Case 32: v = Weekday(Date:=DateSerial(2020, 3, 1), FirstDayOfWeek:=vbMonday)   ' named arguments
        Case 33: v = TypeName(Weekday(DateSerial(2020, 2, 29)))   ' result type
        Case 34: v = TypeName(Weekday(Null))   ' result type for Null
    End Select
    T = "Weekday." & k & "=" & Sh(v)
    Exit Function
fail:
    Dim en As Long
    en = Err.Number
    T = "Weekday." & k & "=ERR" & en
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
