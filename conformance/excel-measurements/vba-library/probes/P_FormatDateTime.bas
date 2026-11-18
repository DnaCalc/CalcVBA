Attribute VB_Name = "P_FormatDateTime"
Option Explicit
' FormatDateTime (Strings; MS-VBAL 6.1.2.11.1.x). Results depend on regional settings: locale.
Private Const N As Long = 25

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String, d As Date
    On Error GoTo fail
    d = DateSerial(2024, 3, 5) + TimeSerial(14, 7, 9)
    Select Case k
        Case 1: v = FormatDateTime(d)                                ' locale: default is vbGeneralDate
        Case 2: v = FormatDateTime(d, vbGeneralDate)                 ' locale
        Case 3: v = FormatDateTime(d, vbLongDate)                    ' locale
        Case 4: v = FormatDateTime(d, vbShortDate)                   ' locale
        Case 5: v = FormatDateTime(d, vbLongTime)                    ' locale
        Case 6: v = FormatDateTime(d, vbShortTime)                   ' locale
        Case 7: v = FormatDateTime(DateSerial(2024, 3, 5), vbGeneralDate)   ' locale: date only
        Case 8: v = FormatDateTime(DateSerial(2024, 3, 5), vbLongDate)      ' locale
        Case 9: v = FormatDateTime(DateSerial(2024, 3, 5), vbShortDate)     ' locale
        Case 10: v = FormatDateTime(DateSerial(2024, 3, 5), vbLongTime)     ' locale: midnight
        Case 11: v = FormatDateTime(DateSerial(2024, 3, 5), vbShortTime)    ' locale
        Case 12: v = FormatDateTime(TimeSerial(14, 7, 9), vbGeneralDate)    ' locale: time only
        Case 13: v = FormatDateTime(TimeSerial(14, 7, 9), vbLongDate)       ' locale: 1899-12-30
        Case 14: v = FormatDateTime(TimeSerial(14, 7, 9), vbShortDate)      ' locale
        Case 15: v = FormatDateTime(TimeSerial(14, 7, 9), vbLongTime)       ' locale
        Case 16: v = FormatDateTime(TimeSerial(14, 7, 9), vbShortTime)      ' locale
        Case 17: v = FormatDateTime("2024/03/05 14:07:09", vbGeneralDate)   ' locale: date string
        Case 18: v = FormatDateTime("abc")                           ' error 13
        Case 19: v = FormatDateTime(Null)                            ' error 94
        Case 20: v = FormatDateTime(Empty)                           ' locale
        Case 21: v = FormatDateTime(45000.5, vbLongDate)             ' locale: number as serial
        Case 22: v = FormatDateTime(d, 5)                            ' error 5: bad NamedFormat
        Case 23: v = FormatDateTime(d, -1)                           ' error 5
        Case 24: v = FormatDateTime(Expression:=d, NamedFormat:=vbShortTime)   ' locale: named
        Case 25: v = FormatDateTime(DateSerial(100, 1, 1), vbLongDate)         ' locale: early year
    End Select
    T = "FormatDateTime." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "FormatDateTime." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
