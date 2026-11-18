Attribute VB_Name = "P_DateDiff"
Option Explicit
' DateDiff (VBA.DateTime; MS-VBAL 6.1.2.x).
Private Const N As Long = 51

Private Function T(ByVal k As Long) As String
    Dim v As Variant, a As Variant, b As Variant, c As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = DateDiff("yyyy", DateSerial(2019, 12, 31), DateSerial(2020, 1, 1))   ' year boundary counts one
        Case 2: v = DateDiff("yyyy", DateSerial(2020, 1, 1), DateSerial(2020, 12, 31))   ' same year is zero
        Case 3: v = DateDiff("q", DateSerial(2020, 3, 31), DateSerial(2020, 4, 1))   ' quarter boundary
        Case 4: v = DateDiff("m", DateSerial(2020, 1, 31), DateSerial(2020, 2, 1))   ' month boundary
        Case 5: v = DateDiff("m", DateSerial(2020, 1, 1), DateSerial(2020, 1, 31))   ' same month is zero
        Case 6: v = DateDiff("y", DateSerial(2020, 1, 1), DateSerial(2020, 3, 1))   ' day of year counts days
        Case 7: v = DateDiff("d", DateSerial(2020, 2, 28), DateSerial(2020, 3, 1))   ' days across Feb 29
        Case 8: v = DateDiff("d", DateSerial(2020, 3, 1), DateSerial(2020, 2, 28))   ' negative result
        Case 9: v = DateDiff("d", #1/1/2020 11:59:59 PM#, #1/2/2020 12:00:01 AM#)   ' two seconds apart across midnight
        Case 10: v = DateDiff("w", DateSerial(2020, 1, 1), DateSerial(2020, 1, 15))   ' weekday counts weeks of the start day
        Case 11: v = DateDiff("ww", DateSerial(2020, 2, 29), DateSerial(2020, 3, 1))   ' Sat to Sun, Sunday-first
        Case 12: v = DateDiff("ww", DateSerial(2020, 2, 29), DateSerial(2020, 3, 1), vbMonday)   ' Sat to Sun, Monday-first
        Case 13: v = DateDiff("ww", DateSerial(2020, 3, 1), DateSerial(2020, 3, 2), vbMonday)   ' Sun to Mon, Monday-first
        Case 14: v = DateDiff("ww", DateSerial(2020, 3, 1), DateSerial(2020, 3, 2), vbSunday)   ' Sun to Mon, Sunday-first
        Case 15: v = DateDiff("ww", DateSerial(2020, 3, 1), DateSerial(2020, 3, 2), vbSaturday)   ' Sun to Mon, Saturday-first
        Case 16: v = DateDiff("ww", DateSerial(2020, 3, 1), DateSerial(2020, 3, 8), vbUseSystemDayOfWeek)   ' system first day ' locale
        Case 17: v = DateDiff("w", DateSerial(2020, 3, 1), DateSerial(2020, 3, 2), vbMonday)   ' w ignores first day of week
        Case 18: v = DateDiff("h", #1/1/2020 11:30:00 PM#, #1/2/2020 12:30:00 AM#)   ' hours
        Case 19: v = DateDiff("h", #1/1/2020 10:59:59 AM#, #1/1/2020 11:00:00 AM#)   ' hour boundary counts one
        Case 20: v = DateDiff("n", #1/1/2020 10:00:59 AM#, #1/1/2020 10:01:00 AM#)   ' minute boundary counts one
        Case 21: v = DateDiff("s", #1/1/2020 10:00:00 AM#, #1/1/2020 10:00:30 AM#)   ' seconds
        Case 22: v = DateDiff("s", #1/1/2020 10:00:30 AM#, #1/1/2020 10:00:00 AM#)   ' negative seconds
        Case 23: v = DateDiff("d", #12/29/1899 6:00:00 AM#, #12/30/1899 6:00:00 AM#)   ' negative serial with time parts
        Case 24: v = DateDiff("h", #12/29/1899 6:00:00 AM#, #12/30/1899 6:00:00 AM#)   ' hours across serial zero
        Case 25: v = DateDiff("yyyy", DateSerial(1800, 1, 1), DateSerial(1900, 1, 1))   ' far before 1899
        Case 26: v = DateDiff("yyyy", DateSerial(100, 1, 1), DateSerial(9999, 12, 31))   ' full range
        Case 27: v = DateDiff("d", DateSerial(100, 1, 1), DateSerial(9999, 12, 31))   ' days across the full range
        Case 28: v = DateDiff("s", DateSerial(100, 1, 1), DateSerial(9999, 12, 31))   ' seconds across full range: Long overflow, error 6
        Case 29: v = DateDiff("ww", DateSerial(2018, 12, 31), DateSerial(2019, 1, 1), vbSunday, vbFirstJan1)   ' week-of-year rules affect ww
        Case 30: v = DateDiff("ww", DateSerial(2018, 12, 30), DateSerial(2019, 1, 6), vbSunday, vbFirstFourDays)   ' four-day rule
        Case 31: v = DateDiff("ww", DateSerial(2018, 12, 30), DateSerial(2019, 1, 6), vbSunday, vbFirstFullWeek)   ' full-week rule
        Case 32: v = DateDiff("yyyy", DateSerial(2018, 12, 31), DateSerial(2019, 1, 1), vbMonday, vbFirstFourDays)   ' year ignores week rules
        Case 33: v = DateDiff("x", DateSerial(2020, 1, 1), DateSerial(2020, 1, 2))   ' bad interval: error 5
        Case 34: v = DateDiff("", DateSerial(2020, 1, 1), DateSerial(2020, 1, 2))   ' empty interval: error 5
        Case 35: v = DateDiff("D", DateSerial(2020, 1, 1), DateSerial(2020, 1, 2))   ' upper-case interval
        Case 36: a = Null: v = DateDiff("d", a, DateSerial(2020, 1, 2))   ' Null first date propagates
        Case 37: a = Null: v = DateDiff("d", DateSerial(2020, 1, 2), a)   ' Null second date propagates
        Case 38: a = Null: v = DateDiff(a, DateSerial(2020, 1, 1), DateSerial(2020, 1, 2))   ' Null interval: error 94
        Case 39: a = Null: v = DateDiff("d", DateSerial(2020, 1, 1), DateSerial(2020, 1, 2), a)   ' Null first day of week: error 94
        Case 40: v = DateDiff("d", DateSerial(2020, 1, 1), DateSerial(2020, 1, 2), 8)   ' first day of week out of range: error 5
        Case 41: v = DateDiff("d", DateSerial(2020, 1, 1), DateSerial(2020, 1, 2), vbSunday, 4)   ' first week of year out of range: error 5
        Case 42: v = DateDiff("d", "1/1/2020", "1/31/2020")   ' string dates ' locale
        Case 43: v = DateDiff("d", "2020/01/01", "2020/01/31")   ' string dates, year first ' locale
        Case 44: v = DateDiff("d", "abc", DateSerial(2020, 1, 1))   ' invalid string date: error 13
        Case 45: v = DateDiff("d", Empty, DateSerial(1899, 12, 31))   ' Empty date is serial 0
        Case 46: v = DateDiff("d", 43831#, 43832#)   ' numeric date serials
        Case 47: v = DateDiff("d", 0.9, 1.1)   ' fractional serials in the same day span
        Case 48: v = DateDiff(Interval:="m", Date1:=DateSerial(2020, 1, 1), Date2:=DateSerial(2020, 6, 1), FirstDayOfWeek:=vbMonday, FirstWeekOfYear:=vbFirstFourDays)   ' named arguments
        Case 49: v = TypeName(DateDiff("d", DateSerial(2020, 1, 1), DateSerial(2020, 1, 2)))   ' result type is Long
        Case 50: v = TypeName(DateDiff("d", Null, DateSerial(2020, 1, 2)))   ' result type with Null
        Case 51: v = DateDiff("d", DateSerial(2020, 1, 1), DateSerial(2020, 1, 1))   ' identical dates
    End Select
    T = "DateDiff." & k & "=" & Sh(v)
    Exit Function
fail:
    Dim en As Long
    en = Err.Number
    T = "DateDiff." & k & "=ERR" & en
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
