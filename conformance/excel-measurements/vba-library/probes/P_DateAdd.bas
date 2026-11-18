Attribute VB_Name = "P_DateAdd"
Option Explicit
' DateAdd (VBA.DateTime; MS-VBAL 6.1.2.x).
Private Const N As Long = 50

Private Function T(ByVal k As Long) As String
    Dim v As Variant, a As Variant, b As Variant, c As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = DateAdd("yyyy", 1, DateSerial(2020, 1, 31))   ' add a year
        Case 2: v = DateAdd("q", 1, DateSerial(2020, 1, 31))   ' add a quarter, month-end clamp to Apr 30
        Case 3: v = DateAdd("m", 1, DateSerial(2020, 1, 31))   ' Jan 31 + 1 month in a leap year clamps to Feb 29
        Case 4: v = DateAdd("m", 1, DateSerial(2021, 1, 31))   ' Jan 31 + 1 month clamps to Feb 28
        Case 5: v = DateAdd("m", -1, DateSerial(2020, 3, 31))   ' month back clamps to Feb 29
        Case 6: v = DateAdd("yyyy", 1, DateSerial(2020, 2, 29))   ' leap day plus a year clamps to Feb 28
        Case 7: v = DateAdd("yyyy", 4, DateSerial(2020, 2, 29))   ' leap day plus four years stays Feb 29
        Case 8: v = DateAdd("y", 10, DateSerial(2020, 12, 25))   ' day of year adds days
        Case 9: v = DateAdd("d", -1, DateSerial(2020, 3, 1))   ' day back across Feb 29
        Case 10: v = DateAdd("w", 3, DateSerial(2020, 2, 27))   ' weekday adds days
        Case 11: v = DateAdd("ww", 2, DateSerial(2020, 2, 27))   ' week adds seven days
        Case 12: v = DateAdd("h", 25, #1/1/2020 10:30:00 AM#)   ' hours cross a day
        Case 13: v = DateAdd("n", 90, #1/1/2020 10:30:00 AM#)   ' minutes
        Case 14: v = DateAdd("s", -90, #1/1/2020 10:30:00 AM#)   ' negative seconds
        Case 15: v = DateAdd("d", 1.5, DateSerial(2020, 1, 1))   ' fractional number rounds to even (2)
        Case 16: v = DateAdd("d", 2.5, DateSerial(2020, 1, 1))   ' 2.5 rounds to even (2)
        Case 17: v = DateAdd("d", -1.5, DateSerial(2020, 1, 10))   ' -1.5 rounds to even (-2)
        Case 18: v = DateAdd("m", 1.5, DateSerial(2020, 1, 31))   ' fractional months
        Case 19: v = DateAdd("h", 1.5, #1/1/2020 10:00:00 AM#)   ' fractional hours round
        Case 20: v = DateAdd("s", 0.4, #1/1/2020 10:00:00 AM#)   ' fractional seconds below half
        Case 21: v = DateAdd("n", 30, #12/29/1899 6:00:00 AM#)   ' negative serial with time part
        Case 22: v = DateAdd("h", 5, #12/29/1899 6:00:00 AM#)   ' crossing serial zero from negative
        Case 23: v = DateAdd("d", 1, #12/30/1899#)   ' serial zero plus a day
        Case 24: v = DateAdd("d", -1, #12/30/1899 6:00:00 AM#)   ' from serial zero with a time part
        Case 25: v = DateAdd("yyyy", -1, DateSerial(100, 1, 1))   ' below year 100: error 5
        Case 26: v = DateAdd("yyyy", 1, DateSerial(9999, 1, 1))   ' above year 9999: error 5
        Case 27: v = DateAdd("d", 1, DateSerial(9999, 12, 31))   ' past the last day: error 5
        Case 28: v = DateAdd("d", -1, DateSerial(100, 1, 1))   ' before the first day: error 5
        Case 29: v = DateAdd("yyyy", 7900, DateSerial(2000, 1, 1))   ' lands on 9900 (ok)
        Case 30: v = DateAdd("x", 1, DateSerial(2020, 1, 1))   ' unknown interval: error 5
        Case 31: v = DateAdd("", 1, DateSerial(2020, 1, 1))   ' empty interval: error 5
        Case 32: v = DateAdd("YYYY", 1, DateSerial(2020, 1, 1))   ' upper-case interval
        Case 33: v = DateAdd(" d", 1, DateSerial(2020, 1, 1))   ' interval with a leading space: error 5
        Case 34: v = DateAdd("dd", 1, DateSerial(2020, 1, 1))   ' two-letter interval: error 5
        Case 35: v = DateAdd("d", 1, Null)   ' Null date propagates
        Case 36: a = Null: v = DateAdd(a, 1, DateSerial(2020, 1, 1))   ' Null interval: error 94
        Case 37: a = Null: v = DateAdd("d", a, DateSerial(2020, 1, 1))   ' Null number: error 94
        Case 38: v = DateAdd("d", 1, "1/31/2020")   ' string date ' locale
        Case 39: v = DateAdd("d", 1, "2020/01/31")   ' string date, year first ' locale
        Case 40: v = DateAdd("d", 1, "abc")   ' invalid date string: error 13
        Case 41: v = DateAdd("d", "2", DateSerial(2020, 1, 1))   ' numeric string for Number
        Case 42: v = DateAdd("d", "x", DateSerial(2020, 1, 1))   ' non-numeric Number: error 13
        Case 43: v = DateAdd("d", 1, 43831#)   ' numeric date serial
        Case 44: v = DateAdd("d", 1, Empty)   ' Empty date is serial 0
        Case 45: v = DateAdd("d", True, DateSerial(2020, 1, 1))   ' Boolean Number is -1
        Case 46: v = DateAdd("d", 1E+10, DateSerial(2020, 1, 1))   ' huge number: error 6 or 5
        Case 47: v = DateAdd(Interval:="m", Number:=2, Date:=DateSerial(2020, 1, 31))   ' named arguments
        Case 48: v = TypeName(DateAdd("d", 1, DateSerial(2020, 1, 1)))   ' result type is Date
        Case 49: v = TypeName(DateAdd("d", 1, "1/1/2020"))   ' result type with a string date ' locale
        Case 50: v = TypeName(DateAdd("d", 1, Null))   ' result type with Null
    End Select
    T = "DateAdd." & k & "=" & Sh(v)
    Exit Function
fail:
    Dim en As Long
    en = Err.Number
    T = "DateAdd." & k & "=ERR" & en
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
