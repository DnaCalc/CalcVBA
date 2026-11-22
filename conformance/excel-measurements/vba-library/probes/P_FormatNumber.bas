Attribute VB_Name = "P_FormatNumber"
Option Explicit
' FormatNumber (Strings; MS-VBAL 6.1.2.11.1.x). All results depend on regional settings: locale.
Private Const N As Long = 36

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = FormatNumber(1234.5678)                          ' locale: defaults
        Case 2: v = FormatNumber(1234.5678, 0)                       ' locale
        Case 3: v = FormatNumber(1234.5678, 5)                       ' locale
        Case 4: v = FormatNumber(1234.5678, -1)                      ' locale: -1 is default digits
        Case 5: v = FormatNumber(0.5, , vbFalse)                     ' locale: no leading digit
        Case 6: v = FormatNumber(0.5, , vbTrue)                      ' locale
        Case 7: v = FormatNumber(0.5, , vbUseDefault)                ' locale
        Case 8: v = FormatNumber(-0.5, 2, vbFalse)                   ' locale: negative, no leading digit
        Case 9: v = FormatNumber(-5.5, , , vbTrue)                   ' locale: parentheses
        Case 10: v = FormatNumber(-5.5, , , vbFalse)                 ' locale
        Case 11: v = FormatNumber(-5.5, , , vbUseDefault)            ' locale
        Case 12: v = FormatNumber(1234567.891, 2, , , vbTrue)        ' locale: grouping
        Case 13: v = FormatNumber(1234567.891, 2, , , vbFalse)       ' locale
        Case 14: v = FormatNumber(1234567.891, 2, , , vbUseDefault)  ' locale
        Case 15: v = FormatNumber(Null)                              ' error 94
        Case 16: v = FormatNumber("abc")                             ' error 13
        Case 17: v = FormatNumber("1234.5")                          ' locale: numeric string
        Case 18: v = FormatNumber(Empty)                             ' locale
        Case 19: v = FormatNumber(True)                              ' locale
        Case 20: v = FormatNumber(0.5, 0)                            ' locale: rounding at .5
        Case 21: v = FormatNumber(1.5, 0)                            ' locale
        Case 22: v = FormatNumber(2.5, 0)                            ' locale
        Case 23: v = FormatNumber(2.675, 2)                          ' locale
        Case 24: v = FormatNumber(CCur(1234.5678), 3)                ' locale: Currency input
        Case 25: v = FormatNumber(CDec("1234.565"), 2)               ' locale: Decimal input
        Case 26: v = FormatNumber(1E+20, 2)                          ' locale: large
        Case 27: v = FormatNumber(0, 2, vbFalse)                     ' locale: zero
        Case 28: v = FormatNumber(Expression:=-1234.5, NumDigitsAfterDecimal:=1, IncludeLeadingDigit:=vbTrue, UseParensForNegativeNumbers:=vbTrue, GroupDigits:=vbTrue)  ' locale: named
        Case 29: v = FormatNumber(5, 100)                            ' 100 fractional digits
        Case 30: v = FormatNumber(5, -2)                             ' error 5
        Case 31: v = FormatNumber(5, 101)
        Case 32: v = FormatNumber(5, 127)
        Case 33: v = FormatNumber(5, 128)
        Case 34: v = FormatNumber(5, 255)
        Case 35: v = FormatNumber(5, 256)
        Case 36: v = FormatNumber(Empty, 256, 2)
    End Select
    T = "FormatNumber." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "FormatNumber." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
