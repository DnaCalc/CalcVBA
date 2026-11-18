Attribute VB_Name = "P_DateValue"
Option Explicit
' DateValue (VBA.DateTime; MS-VBAL 6.1.2.x).
Private Const N As Long = 27

Private Function T(ByVal k As Long) As String
    Dim v As Variant, a As Variant, b As Variant, c As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = DateValue("2020/02/29")   ' year first ' locale
        Case 2: v = DateValue("February 29, 2020")   ' month name, day, year
        Case 3: v = DateValue("29 Feb 2020")   ' day, month name, year
        Case 4: v = DateValue("Feb 29 2020 10:30 PM")   ' time part is dropped
        Case 5: v = DateValue("2020-02-29")   ' ISO form ' locale
        Case 6: v = DateValue("1/31/2020")   ' month first ' locale
        Case 7: v = DateValue("31/1/2020")   ' day first ' locale
        Case 8: v = DateValue("13/13/2020")   ' no valid reading: error 13
        Case 9: v = DateValue("2020-02-30")   ' invalid day: error 13
        Case 10: v = DateValue("2/29/1900")   ' invalid leap day: error 13
        Case 11: v = DateValue("1/1/29")   ' two-digit year 29 ' locale
        Case 12: v = DateValue("1/1/30")   ' two-digit year 30 ' locale
        Case 13: v = DateValue("12/31/9999")   ' upper limit ' locale
        Case 14: v = DateValue("1/1/100")   ' lower limit ' locale
        Case 15: v = DateValue("12:30:00")   ' time-only text ' locale
        Case 16: v = DateValue("")   ' empty string: error 13
        Case 17: v = DateValue("abc")   ' invalid text: error 13
        Case 18: v = DateValue(Null)   ' Null argument
        Case 19: v = DateValue(Empty)   ' Empty argument
        Case 20: v = DateValue(123)   ' number converted to text: error 13 or a date
        Case 21: a = DateSerial(2020, 2, 29) + TimeSerial(10, 0, 0): v = DateValue(a)   ' Date argument converted to text and back
        Case 22: v = DateValue("Jan 5") = DateSerial(Year(Date), 1, 5)   ' year defaults to current year
        Case 23: v = TypeName(DateValue("2020/02/29"))   ' result type ' locale
        Case 24: v = DateValue(Date:="2020/01/02")   ' named argument ' locale
        Case 25: v = DateValue("  2020/02/29  ")   ' surrounding spaces ' locale
        Case 26: v = DateValue("1 1 2020")   ' space separators ' locale
        Case 27: v = DateValue("Sunday, February 2, 2020")   ' weekday name in text
    End Select
    T = "DateValue." & k & "=" & Sh(v)
    Exit Function
fail:
    Dim en As Long
    en = Err.Number
    T = "DateValue." & k & "=ERR" & en
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
