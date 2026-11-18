Attribute VB_Name = "P_Year"
Option Explicit
' Year (VBA.DateTime; MS-VBAL 6.1.2.x).
Private Const N As Long = 26

Private Function T(ByVal k As Long) As String
    Dim v As Variant, a As Variant, b As Variant, c As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Year(DateSerial(2020, 2, 29))   ' typical leap day
        Case 2: v = Year(DateSerial(9999, 12, 31))   ' upper limit
        Case 3: v = Year(DateSerial(100, 1, 1))   ' lower limit
        Case 4: v = Year(0)   ' serial 0 is 1899
        Case 5: v = Year(-1)   ' negative serial
        Case 6: v = Year(-0.5)   ' negative fraction
        Case 7: v = Year(True)   ' Boolean -1
        Case 8: v = Year(Null)   ' Null propagates
        Case 9: v = Year(Empty)   ' Empty is serial 0
        Case 10: v = Year("2020/02/29")   ' string date ' locale
        Case 11: v = Year("February 29, 2020")   ' string date with month name
        Case 12: v = Year("1/1/29")   ' two-digit year 29 in text ' locale
        Case 13: v = Year("1/1/30")   ' two-digit year 30 in text ' locale
        Case 14: v = Year("abc")   ' invalid string: error 13
        Case 15: v = Year("")   ' empty string: error 13
        Case 16: v = Year("2020")   ' numeric string
        Case 17: v = Year(2958465)   ' serial of 9999-12-31
        Case 18: v = Year(2958466)   ' beyond upper limit: error 5
        Case 19: v = Year(-657434)   ' serial of 0100-01-01
        Case 20: v = Year(-657435)   ' below lower limit: error 5
        Case 21: v = Year(Date:=DateSerial(2021, 3, 15))   ' named argument
        Case 22: v = TypeName(Year(DateSerial(2020, 2, 29)))   ' result type
        Case 23: v = TypeName(Year(Null))   ' result type for Null
        Case 24: v = Year(#1/1/1900#)   ' date literal
        Case 25: v = Year(36526)   ' serial of 2000-01-01
        Case 26: v = Year(43830.9999)   ' end of 2019
    End Select
    T = "Year." & k & "=" & Sh(v)
    Exit Function
fail:
    Dim en As Long
    en = Err.Number
    T = "Year." & k & "=ERR" & en
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
