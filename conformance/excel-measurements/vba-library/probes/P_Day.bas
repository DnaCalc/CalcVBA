Attribute VB_Name = "P_Day"
Option Explicit
' Day (VBA.DateTime; MS-VBAL 6.1.2.x).
Private Const N As Long = 29

Private Function T(ByVal k As Long) As String
    Dim v As Variant, a As Variant, b As Variant, c As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Day(DateSerial(2020, 2, 29))   ' typical leap day
        Case 2: v = Day(DateSerial(9999, 12, 31))   ' upper limit
        Case 3: v = Day(DateSerial(100, 1, 1))   ' lower limit
        Case 4: v = Day(0)   ' serial 0 is 1899-12-30
        Case 5: v = Day(-1)   ' negative serial is 1899-12-29
        Case 6: v = Day(-0.5)   ' negative fraction
        Case 7: v = Day(1.9999)   ' fraction before next day
        Case 8: v = Day(True)   ' Boolean -1
        Case 9: v = Day(Null)   ' Null propagates
        Case 10: v = Day(Empty)   ' Empty is serial 0
        Case 11: v = Day("2020/02/29")   ' string date ' locale
        Case 12: v = Day("February 29, 2020")   ' string date with month name
        Case 13: v = Day("1/31/2020")   ' string date, month first ' locale
        Case 14: v = Day("abc")   ' invalid string: error 13
        Case 15: v = Day("")   ' empty string: error 13
        Case 16: v = Day("31")   ' numeric string
        Case 17: v = Day(2958465)   ' serial of 9999-12-31
        Case 18: v = Day(2958466)   ' beyond upper limit: error 5
        Case 19: v = Day(-657434)   ' serial of 0100-01-01
        Case 20: v = Day(-657435)   ' below lower limit: error 5
        Case 21: v = Day(1E+30)   ' huge number: error 5 or 6
        Case 22: v = Day(DateSerial(2020, 1, 31) + 0.9999)   ' date with time part
        Case 23: v = Day(Date:=DateSerial(2021, 3, 15))   ' named argument
        Case 24: v = TypeName(Day(DateSerial(2020, 2, 29)))   ' result type
        Case 25: v = TypeName(Day(Null))   ' result type for Null
        Case 26: v = Day(#12/31/1899#)   ' date literal
        Case 27: v = Day(30.7)   ' serial 30.7 fraction ignored
        Case 28: v = Day(CDate("12/30/1899 6:00 PM"))   ' serial zero with time ' locale
        Case 29: v = Day(CVErr(2007))   ' error value argument: error 13
    End Select
    T = "Day." & k & "=" & Sh(v)
    Exit Function
fail:
    Dim en As Long
    en = Err.Number
    T = "Day." & k & "=ERR" & en
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
