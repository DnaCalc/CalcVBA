Attribute VB_Name = "P_DatePart"
Option Explicit
' DatePart (VBA.DateTime; MS-VBAL 6.1.2.x).
Private Const N As Long = 46

Private Function T(ByVal k As Long) As String
    Dim v As Variant, a As Variant, b As Variant, c As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = DatePart("yyyy", DateSerial(2020, 2, 29))   ' year
        Case 2: v = DatePart("q", DateSerial(2020, 2, 29))   ' quarter
        Case 3: v = DatePart("q", DateSerial(2020, 12, 31))   ' last quarter
        Case 4: v = DatePart("m", DateSerial(2020, 2, 29))   ' month
        Case 5: v = DatePart("y", DateSerial(2020, 12, 31))   ' day of year in a leap year
        Case 6: v = DatePart("y", DateSerial(2021, 12, 31))   ' day of year in a normal year
        Case 7: v = DatePart("d", DateSerial(2020, 2, 29))   ' day
        Case 8: v = DatePart("w", DateSerial(2020, 2, 29))   ' weekday, Sunday-first
        Case 9: v = DatePart("w", DateSerial(2020, 2, 29), vbMonday)   ' weekday, Monday-first
        Case 10: v = DatePart("w", DateSerial(2020, 3, 1), vbMonday)   ' Sunday under Monday-first is 7
        Case 11: v = DatePart("ww", DateSerial(2020, 1, 1))   ' week of year, default rules
        Case 12: v = DatePart("ww", DateSerial(2018, 12, 31))   ' late December, default rules
        Case 13: v = DatePart("ww", DateSerial(2021, 1, 1))   ' Jan 1 2021 (Friday), default rules
        Case 14: v = DatePart("ww", DateSerial(2021, 1, 1), vbMonday, vbFirstFourDays)   ' ISO-like rules
        Case 15: v = DatePart("ww", DateSerial(2021, 1, 3), vbMonday, vbFirstFourDays)   ' ISO week 53 of 2020
        Case 16: v = DatePart("ww", DateSerial(2021, 1, 4), vbMonday, vbFirstFourDays)   ' ISO week 1
        Case 17: v = DatePart("ww", DateSerial(2021, 1, 1), vbSunday, vbFirstFullWeek)   ' full-week rule
        Case 18: v = DatePart("ww", DateSerial(2021, 1, 10), vbSunday, vbFirstFullWeek)   ' full-week rule, second week
        Case 19: v = DatePart("ww", DateSerial(2021, 1, 1), vbSunday, vbFirstJan1)   ' week containing Jan 1
        Case 20: v = DatePart("ww", DateSerial(2021, 12, 31), vbSunday, vbFirstJan1)   ' end of year
        Case 21: v = DatePart("ww", DateSerial(2021, 1, 1), vbUseSystemDayOfWeek, vbUseSystem)   ' system rules ' locale
        Case 22: v = DatePart("h", #1/1/2020 3:04:05 PM#)   ' hour
        Case 23: v = DatePart("n", #1/1/2020 3:04:05 PM#)   ' minute
        Case 24: v = DatePart("s", #1/1/2020 3:04:05 PM#)   ' second
        Case 25: v = DatePart("h", #12/29/1899 6:30:00 AM#)   ' negative serial with time part
        Case 26: v = DatePart("d", #12/29/1899 6:30:00 AM#)   ' day of a negative serial
        Case 27: v = DatePart("yyyy", DateSerial(100, 1, 1))   ' lower limit
        Case 28: v = DatePart("yyyy", DateSerial(9999, 12, 31))   ' upper limit
        Case 29: v = DatePart("ww", DateSerial(100, 1, 1))   ' week at the lower limit
        Case 30: v = DatePart("x", DateSerial(2020, 1, 1))   ' bad interval: error 5
        Case 31: v = DatePart("", DateSerial(2020, 1, 1))   ' empty interval: error 5
        Case 32: v = DatePart("M", DateSerial(2020, 2, 29))   ' upper-case interval
        Case 33: v = DatePart("Q", DateSerial(2020, 2, 29))   ' upper-case quarter
        Case 34: v = DatePart("m", Null)   ' Null date propagates
        Case 35: a = Null: v = DatePart(a, DateSerial(2020, 1, 1))   ' Null interval: error 94
        Case 36: a = Null: v = DatePart("ww", DateSerial(2020, 1, 1), a)   ' Null first day of week: error 94
        Case 37: v = DatePart("ww", DateSerial(2020, 1, 1), 8)   ' first day of week out of range: error 5
        Case 38: v = DatePart("ww", DateSerial(2020, 1, 1), vbSunday, 4)   ' first week of year out of range: error 5
        Case 39: v = DatePart("m", "1/31/2020")   ' string date ' locale
        Case 40: v = DatePart("d", "2020/01/31")   ' string date, year first ' locale
        Case 41: v = DatePart("m", "abc")   ' invalid string date: error 13
        Case 42: v = DatePart("m", Empty)   ' Empty is serial 0
        Case 43: v = DatePart("m", 43831#)   ' numeric date serial
        Case 44: v = DatePart(Interval:="w", Date:=DateSerial(2020, 3, 1), FirstDayOfWeek:=vbMonday, FirstWeekOfYear:=vbFirstFourDays)   ' named arguments
        Case 45: v = TypeName(DatePart("m", DateSerial(2020, 2, 29)))   ' result type
        Case 46: v = TypeName(DatePart("m", Null))   ' result type with Null
    End Select
    T = "DatePart." & k & "=" & Sh(v)
    Exit Function
fail:
    Dim en As Long
    en = Err.Number
    T = "DatePart." & k & "=ERR" & en
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
