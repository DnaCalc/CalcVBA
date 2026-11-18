Attribute VB_Name = "P_CDate"
Option Explicit
' CDate (VBA.Conversion; MS-VBAL 6.1.2.x).
Private Const N As Long = 47

Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, c As Currency, dt As Date, ds As Double, sg As Single
    Dim l As Long, i As Integer, by As Byte, b As Boolean, s As String, o As Object
    On Error GoTo fail
    Select Case k
        Case 1: v = CDate(0)   ' epoch 1899/12/30
        Case 2: v = CDate(1)
        Case 3: v = CDate(1.5)
        Case 4: v = CDate(-1.5)   ' before epoch, time part
        Case 5: v = CDate(43831)
        Case 6: v = CDate(43831.75)
        Case 7: v = CDate(2958465.99999)   ' near max
        Case 8: v = CDate(2958466)   ' overflow
        Case 9: v = CDate(-657434)   ' min 100/01/01
        Case 10: v = CDate(-657435)   ' overflow
        Case 11: v = CDate(2958465.999999999)   ' sub-second rounding
        Case 12: v = CDate(#1/2/2020#)   ' date literal
        Case 13: v = CDate(True)   ' -1
        Case 14: v = CDate(False)
        Case 15: v = CDate(Null)   ' 94
        Case 16: v = CDate(Empty)
        Case 17: v = CDate(CVErr(5))   ' 13
        Case 18: v = CDate("")   ' 13
        Case 19: v = CDate("abc")   ' 13
        Case 20: v = CDate("2020/01/02")   ' ISO-like, locale
        Case 21: v = CDate("1/2/2020")   ' ambiguous, locale
        Case 22: v = CDate("13/1/2020")   ' locale
        Case 23: v = CDate("2020-01-02")   ' locale
        Case 24: v = CDate("2 January 2020")   ' locale
        Case 25: v = CDate("January 2, 2020")   ' locale
        Case 26: v = CDate("Jan 2 2020 3:04 PM")   ' locale
        Case 27: v = CDate("12:30")   ' time only
        Case 28: v = CDate("12:30:45.5")   ' locale
        Case 29: v = CDate("3:04 PM")
        Case 30: v = CDate("25:00")   ' 13
        Case 31: v = CDate("1")   ' numeric string, locale
        Case 32: v = CDate("43831")   ' numeric string
        Case 33: v = CDate("43831.5")
        Case 34: v = CDate("#1/2/2020#")   ' hash delimited string
        Case 35: v = CDate("2/29/2019")   ' invalid day 13
        Case 36: v = CDate("2/29/2020")   ' leap, locale
        Case 37: v = CDate("12/31/9999")   ' max date, locale
        Case 38: v = CDate("1/1/100")   ' locale
        Case 39: v = CDate("1/1/29")   ' 2-digit year cutoff, locale
        Case 40: v = CDate("1/1/30")   ' locale
        Case 41: v = CDate(CCur(1.5))
        Case 42: v = CDate(CDec(2.25))
        Case 43: v = CDate(CSng(1.5))
        Case 44: v = CDate(CByte(5))
        Case 45: v = CDate(" 1/2/2020 ")   ' spaces, locale
        Case 46: v = CDate("Monday")   ' locale
        Case 47: v = CDate(#1/2/2020 3:04:05 PM#)   ' Date passthrough
    End Select
    T = "CDate." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "CDate." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
