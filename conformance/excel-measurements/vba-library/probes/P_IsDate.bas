Attribute VB_Name = "P_IsDate"
Option Explicit
' IsDate (Information; string parsing is locale dependent, marked locale).
Private Const N As Long = 45

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    Dim va As Variant, dt As Date, c As Collection
    On Error GoTo fail
    Select Case k
        Case 1: dt = DateSerial(2020, 1, 2): v = IsDate(dt)   ' Date value
        Case 2: v = IsDate(#1/2/2020#)                        ' Date literal
        Case 3: v = IsDate(DateSerial(1, 1, 1))               ' early date
        Case 4: v = IsDate("2020/01/02")                      ' locale: short date
        Case 5: v = IsDate("1/2/2020")                        ' locale: m/d or d/m
        Case 6: v = IsDate("2020-01-02")                      ' ISO
        Case 7: v = IsDate("2 Jan 2020")
        Case 8: v = IsDate("Jan 5 2020")
        Case 9: v = IsDate("January 5, 2020")
        Case 10: v = IsDate("February 30, 2020")               ' impossible day
        Case 11: v = IsDate("2020-13-45")                      ' impossible month and day
        Case 12: v = IsDate("12:30")                           ' time only
        Case 13: v = IsDate("12:30:45 PM")
        Case 14: v = IsDate("25:00")                           ' impossible time
        Case 15: v = IsDate("13:61")
        Case 16: v = IsDate("2020/01/02 12:30")                ' date and time
        Case 17: v = IsDate("1/1/100")                         ' year 100
        Case 18: v = IsDate("12/31/9999")                      ' last date
        Case 19: v = IsDate("1/1/10000")                       ' year past the end
        Case 20: v = IsDate("abc")
        Case 21: v = IsDate("")
        Case 22: v = IsDate(" ")
        Case 23: v = IsDate("1")                               ' plain number string
        Case 24: v = IsDate("43000")                           ' serial as string
        Case 25: v = IsDate(43000)                             ' number is not a date
        Case 26: v = IsDate(43000#)                            ' Double
        Case 27: v = IsDate(0)
        Case 28: v = IsDate(True)
        Case 29: v = IsDate(CDate(1))                          ' converted to Date
        Case 30: v = IsDate(Null)
        Case 31: v = IsDate(Empty)
        Case 32: v = IsDate(CVErr(5))
        Case 33: v = IsDate(Array(#1/1/2000#))                 ' array
        Case 34: v = IsDate(Array(#1/1/2000#)(0))              ' element
        Case 35: Set c = New Collection: v = IsDate(c)         ' object
        Case 36: v = IsDate("1e3")
        Case 37: v = IsDate("1/2")                             ' month and day only
        Case 38: v = IsDate("Jan")                             ' month name only
        Case 39: v = IsDate("Monday")                          ' weekday name
        Case 40: v = IsDate("1 January")
        Case 41: v = IsDate("#1/2/2020#")                      ' with hash signs
        Case 42: v = IsDate("12.30")
        Case 43: v = IsDate("1/2/20")                          ' two digit year
        Case 44: v = IsDate("2020/02/30")                      ' impossible day, yyyy/mm/dd
        Case 45: v = IsDate(Expression:=#1/1/2000#)           ' named argument
    End Select
    T = "IsDate." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "IsDate." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
