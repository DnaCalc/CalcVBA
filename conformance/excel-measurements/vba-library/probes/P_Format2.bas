Attribute VB_Name = "P_Format2"
Option Explicit
' Format / Format$ part 2: named and user date/time formats (MS-VBAL 6.1.2.11.1.x).
Private Const N As Long = 79

Private Function Dt() As Date
    Dt = DateSerial(2024, 3, 5) + TimeSerial(14, 7, 9)    ' Tuesday 2024-03-05 14:07:09
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String, d As Date, e As Date
    On Error GoTo fail
    d = Dt()
    e = DateSerial(2021, 1, 1) + TimeSerial(0, 5, 3)      ' Friday, early morning
    Select Case k
        Case 1: v = Format(d, "General Date")                 ' locale
        Case 2: v = Format(d, "Long Date")                    ' locale
        Case 3: v = Format(d, "Medium Date")                  ' locale
        Case 4: v = Format(d, "Short Date")                   ' locale
        Case 5: v = Format(d, "Long Time")                    ' locale
        Case 6: v = Format(d, "Medium Time")                  ' locale
        Case 7: v = Format(d, "Short Time")                   ' locale
        Case 8: v = Format(DateSerial(2024, 3, 5), "General Date")   ' locale: date only
        Case 9: v = Format(TimeSerial(14, 7, 9), "General Date")     ' locale: time only
        Case 10: v = Format(d, "d")
        Case 11: v = Format(d, "dd")
        Case 12: v = Format(d, "ddd")                         ' locale
        Case 13: v = Format(d, "dddd")                        ' locale
        Case 14: v = Format(d, "ddddd")                       ' locale
        Case 15: v = Format(d, "dddddd")                      ' locale
        Case 16: v = Format(d, "w")
        Case 17: v = Format(d, "w", vbMonday)
        Case 18: v = Format(d, "w", vbTuesday)
        Case 19: v = Format(e, "ww")
        Case 20: v = Format(e, "ww", vbMonday, vbFirstJan1)
        Case 21: v = Format(e, "ww", vbMonday, vbFirstFourDays)
        Case 22: v = Format(e, "ww", vbMonday, vbFirstFullWeek)
        Case 23: v = Format(e, "ww", vbSunday, vbFirstFullWeek)
        Case 24: v = Format(d, "m")
        Case 25: v = Format(d, "mm")
        Case 26: v = Format(d, "mmm")                         ' locale
        Case 27: v = Format(d, "mmmm")                        ' locale
        Case 28: v = Format(d, "q")
        Case 29: v = Format(d, "y")
        Case 30: v = Format(DateSerial(2024, 12, 31), "y")    ' leap year day 366
        Case 31: v = Format(d, "yy")
        Case 32: v = Format(d, "yyyy")
        Case 33: v = Format(d, "h")
        Case 34: v = Format(d, "hh")
        Case 35: v = Format(d, "h AM/PM")
        Case 36: v = Format(e, "h AM/PM")                     ' midnight hour
        Case 37: v = Format(d, "n")
        Case 38: v = Format(e, "nn")
        Case 39: v = Format(d, "s")
        Case 40: v = Format(e, "ss")
        Case 41: v = Format(d, "ttttt")                       ' locale
        Case 42: v = Format(d, "hh:mm AM/PM")
        Case 43: v = Format(d, "hh:mm am/pm")
        Case 44: v = Format(d, "hh:mm A/P")
        Case 45: v = Format(d, "hh:mm a/p")
        Case 46: v = Format(d, "hh:mm AMPM")                  ' locale: AM/PM designator
        Case 47: v = Format(d, "c")                           ' locale
        Case 48: v = Format(d, "yyyy/mm/dd")                  ' date separator /
        Case 49: v = Format(d, "yyyy/mm/dd hh:nn:ss")         ' locale: separators
        Case 50: v = Format(d, "h:m")                         ' m after h is minute
        Case 51: v = Format(d, "m:s")                         ' m before s is minute
        Case 52: v = Format(d, "mm")                          ' alone is month
        Case 53: v = Format(d, "hh:mm:ss")
        Case 54: v = Format(d, "mm:dd")
        Case 55: v = Format(d, "yyyy-mm-dd\Thh:nn:ss")
        Case 56: v = Format(d, "d mmmm yyyy")                 ' locale
        Case 57: v = Format("2024/03/05", "yyyy-mm-dd")       ' locale: date string parsed
        Case 58: v = Format("abc", "yyyy-mm-dd")              ' non-date string
        Case 59: v = Format(0.5, "hh:nn")                     ' number as serial time
        Case 60: v = Format(45000, "yyyy/mm/dd")
        Case 61: v = Format(TimeSerial(13, 5, 9), "yyyy/mm/dd")   ' time-only gives 1899/12/30
        Case 62: v = Format(DateSerial(100, 1, 1), "yyyy/mm/dd")
        Case 63: v = Format(DateSerial(9999, 12, 31), "yyyy/mm/dd")
        Case 64: v = Format(DateSerial(1900, 3, 1), "dddd")   ' locale
        Case 65: v = Format(d, "dd/mm/yyyy")
        Case 66: v = Format(d, "yyyy ""year"" mm")
        Case 67: v = Format(d, "dddd, mmmm d, yyyy")          ' locale
        Case 68: v = Format(#12:00:00 PM#, "h:mm AM/PM")
        Case 69: v = Format(#12:00:00 AM#, "hh:mm AM/PM")
        Case 70: v = Format(d, "mmmmm")                       ' five m
        Case 71: v = Format(d, "yyy")
        Case 72: v = Format(d, "hh:mm:ss.000")                ' literal period then zeros
        Case 73: v = Format(d, "d/m/yy")
        Case 74: v = Format(d, "mmm-yy")                      ' locale
        Case 75: v = Format(d, "[h]:mm")                      ' brackets are literal
        Case 76: v = Format(d, "a/p")
        Case 77: v = Format(Null, "yyyy")
        Case 78: v = Format$(d, "yyyy")                       ' $ form
        Case 79: v = Format(Expression:=d, Format:="yyyy", FirstDayOfWeek:=vbMonday, FirstWeekOfYear:=vbFirstJan1)
    End Select
    T = "Format2." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Format2." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
