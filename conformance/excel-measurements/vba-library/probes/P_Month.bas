Attribute VB_Name = "P_Month"
Option Explicit
' Month (VBA.DateTime; MS-VBAL 6.1.2.x).
Private Const N As Long = 27

Private Function T(ByVal k As Long) As String
    Dim v As Variant, a As Variant, b As Variant, c As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Month(DateSerial(2020, 2, 29))   ' typical leap day
        Case 2: v = Month(DateSerial(9999, 12, 31))   ' upper limit
        Case 3: v = Month(DateSerial(100, 1, 1))   ' lower limit
        Case 4: v = Month(0)   ' serial 0 is 1899-12-30
        Case 5: v = Month(-1)   ' negative serial
        Case 6: v = Month(-0.5)   ' negative fraction
        Case 7: v = Month(True)   ' Boolean -1
        Case 8: v = Month(Null)   ' Null propagates
        Case 9: v = Month(Empty)   ' Empty is serial 0
        Case 10: v = Month("2020/02/29")   ' string date ' locale
        Case 11: v = Month("February 29, 2020")   ' string date with month name
        Case 12: v = Month("1/31/2020")   ' string date, month first ' locale
        Case 13: v = Month("31/12/2020")   ' string date, day first ' locale
        Case 14: v = Month("abc")   ' invalid string: error 13
        Case 15: v = Month("")   ' empty string: error 13
        Case 16: v = Month("12")   ' numeric string
        Case 17: v = Month(2958465)   ' serial of 9999-12-31
        Case 18: v = Month(2958466)   ' beyond upper limit: error 5
        Case 19: v = Month(-657434)   ' serial of 0100-01-01
        Case 20: v = Month(-657435)   ' below lower limit: error 5
        Case 21: v = Month(Date:=DateSerial(2021, 3, 15))   ' named argument
        Case 22: v = TypeName(Month(DateSerial(2020, 2, 29)))   ' result type
        Case 23: v = TypeName(Month(Null))   ' result type for Null
        Case 24: v = Month(#1/1/1900#)   ' date literal
        Case 25: v = Month(31)   ' serial 31 is 1900-01-30
        Case 26: v = Month(32)   ' serial 32 is 1900-01-31
        Case 27: v = Month(61)   ' serial 61 is 1900-03-01 in VBA
    End Select
    T = "Month." & k & "=" & Sh(v)
    Exit Function
fail:
    Dim en As Long
    en = Err.Number
    T = "Month." & k & "=ERR" & en
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
