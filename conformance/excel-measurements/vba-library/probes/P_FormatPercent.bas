Attribute VB_Name = "P_FormatPercent"
Option Explicit
' FormatPercent (Strings; MS-VBAL 6.1.2.11.1.x). All results depend on regional settings: locale.
Private Const N As Long = 43

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = FormatPercent(1234.5678)                          ' locale: defaults
        Case 2: v = FormatPercent(1234.5678, 0)                       ' locale
        Case 3: v = FormatPercent(1234.5678, 5)                       ' locale
        Case 4: v = FormatPercent(1234.5678, -1)                      ' locale: -1 is default digits
        Case 5: v = FormatPercent(0.5, , vbFalse)                     ' locale: no leading digit
        Case 6: v = FormatPercent(0.5, , vbTrue)                      ' locale
        Case 7: v = FormatPercent(0.5, , vbUseDefault)                ' locale
        Case 8: v = FormatPercent(-0.5, 2, vbFalse)                   ' locale: negative, no leading digit
        Case 9: v = FormatPercent(-5.5, , , vbTrue)                   ' locale: parentheses
        Case 10: v = FormatPercent(-5.5, , , vbFalse)                 ' locale
        Case 11: v = FormatPercent(-5.5, , , vbUseDefault)            ' locale
        Case 12: v = FormatPercent(1234567.891, 2, , , vbTrue)        ' locale: grouping
        Case 13: v = FormatPercent(1234567.891, 2, , , vbFalse)       ' locale
        Case 14: v = FormatPercent(1234567.891, 2, , , vbUseDefault)  ' locale
        Case 15: v = FormatPercent(Null)                              ' error 94
        Case 16: v = FormatPercent("abc")                             ' error 13
        Case 17: v = FormatPercent("1234.5")                          ' locale: numeric string
        Case 18: v = FormatPercent(Empty)                             ' locale
        Case 19: v = FormatPercent(True)                              ' locale
        Case 20: v = FormatPercent(0.5, 0)                            ' locale: rounding at .5
        Case 21: v = FormatPercent(1.5, 0)                            ' locale
        Case 22: v = FormatPercent(2.5, 0)                            ' locale
        Case 23: v = FormatPercent(2.675, 2)                          ' locale
        Case 24: v = FormatPercent(CCur(1234.5678), 3)                ' locale: Currency input
        Case 25: v = FormatPercent(CDec("1234.565"), 2)               ' locale: Decimal input
        Case 26: v = FormatPercent(1E+20, 2)                          ' locale: large
        Case 27: v = FormatPercent(0, 2, vbFalse)                     ' locale: zero
        Case 28: v = FormatPercent(Expression:=-1234.5, NumDigitsAfterDecimal:=1, IncludeLeadingDigit:=vbTrue, UseParensForNegativeNumbers:=vbTrue, GroupDigits:=vbTrue)  ' locale: named
        Case 29: v = FormatPercent(5, 100)                            ' 100 fractional digits
        Case 30: v = FormatPercent(5, -2)                             ' error 5
        Case 31: v = FormatPercent(5, 101)
        Case 32: v = FormatPercent(5, 127)
        Case 33: v = FormatPercent(5, 128)
        Case 34: v = FormatPercent(5, 255)
        Case 35: v = FormatPercent(5, 256)
        Case 36: v = FormatPercent(Empty, 256, 2)
        Case 37: v = FormatPercent(0.005, 256)
        Case 38: v = FormatPercent(5, 257)
        Case 39: v = FormatPercent(0.005, 257)
        Case 40: v = FormatPercent(5, 511)
        Case 41: v = FormatPercent(0.005, 511)
        Case 42: v = FormatPercent(5, 512)
        Case 43: v = FormatPercent(0.005, 512)
    End Select
    T = "FormatPercent." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "FormatPercent." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
