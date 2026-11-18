Attribute VB_Name = "P_DateSerial"
Option Explicit
' DateSerial (VBA.DateTime; MS-VBAL 6.1.2.x).
Private Const N As Long = 38

Private Function T(ByVal k As Long) As String
    Dim v As Variant, a As Variant, b As Variant, c As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = DateSerial(2020, 2, 29)   ' typical leap day
        Case 2: v = DateSerial(2020, 13, 1)   ' month 13 rolls into next year
        Case 3: v = DateSerial(2020, 0, 1)   ' month 0 is December of previous year
        Case 4: v = DateSerial(2020, -1, 1)   ' negative month
        Case 5: v = DateSerial(2020, 1, 0)   ' day 0 is last day of previous month
        Case 6: v = DateSerial(2020, 3, -1)   ' negative day
        Case 7: v = DateSerial(2020, 1, 32)   ' day overflow
        Case 8: v = DateSerial(2020, 14, 63)   ' both overflow
        Case 9: v = DateSerial(29, 1, 1)   ' two-digit year 29 becomes 2029
        Case 10: v = DateSerial(30, 1, 1)   ' two-digit year 30 becomes 1930
        Case 11: v = DateSerial(0, 1, 1)   ' year 0 becomes 2000
        Case 12: v = DateSerial(99, 12, 31)   ' year 99 becomes 1999
        Case 13: v = DateSerial(100, 1, 1)   ' year 100, the lower limit
        Case 14: v = DateSerial(100, 1, 0)   ' just below the lower limit: error 5
        Case 15: v = DateSerial(9999, 12, 31)   ' upper limit
        Case 16: v = DateSerial(9999, 12, 32)   ' just above the upper limit: error 5
        Case 17: v = DateSerial(9999, 13, 1)   ' month overflow past upper limit: error 5
        Case 18: v = DateSerial(10000, 1, 1)   ' year 10000: error 5
        Case 19: v = DateSerial(-1, 1, 1)   ' negative year
        Case 20: v = DateSerial(1899, 12, 30)   ' serial zero
        Case 21: v = DateSerial(1900, 2, 29)   ' non-existent leap day rolls to March 1
        Case 22: v = DateSerial(1900, 1, 1)   ' first day of 1900
        Case 23: a = 32767: v = DateSerial(a, 1, 1)   ' largest Integer year: error 5
        Case 24: a = 32768: v = DateSerial(a, 1, 1)   ' beyond Integer: overflow error 6
        Case 25: a = -32769: v = DateSerial(2020, a, 1)   ' beyond Integer in month: overflow 6
        Case 26: a = 40000: v = DateSerial(2020, 1, a)   ' beyond Integer in day: overflow 6
        Case 27: a = 32767: v = DateSerial(2020, 1, a)   ' largest Integer day: error 5 or a date
        Case 28: a = "2020": b = "2": c = "29": v = DateSerial(a, b, c)   ' numeric strings
        Case 29: a = 2020.5: v = DateSerial(a, 2, 1)   ' fractional year rounds to even (2020)
        Case 30: a = 2021.5: v = DateSerial(a, 2.5, 1.5)   ' fractional parts round to even
        Case 31: v = DateSerial(0, 0, 0)   ' all zero parts
        Case 32: a = Empty: v = DateSerial(a, 1, 1)   ' Empty year is 0, becomes 2000
        Case 33: a = Null: v = DateSerial(a, 1, 1)   ' Null argument: error 94
        Case 34: a = "abc": v = DateSerial(2020, a, 1)   ' non-numeric string: error 13
        Case 35: v = DateSerial(Year:=2021, Month:=3, Day:=15)   ' named arguments
        Case 36: v = DateSerial(Day:=15, Month:=3, Year:=2021)   ' named arguments out of order
        Case 37: v = DateSerial(True, 1, 1)   ' Boolean argument is -1: error 5
        Case 38: v = DateSerial(1800, 6, 15)   ' negative serial result
    End Select
    T = "DateSerial." & k & "=" & Sh(v)
    Exit Function
fail:
    Dim en As Long
    en = Err.Number
    T = "DateSerial." & k & "=ERR" & en
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
