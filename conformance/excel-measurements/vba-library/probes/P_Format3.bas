Attribute VB_Name = "P_Format3"
Option Explicit
' Format / Format$ part 3: string formats, input types and argument errors (MS-VBAL 6.1.2.11.1.x).
Private Const N As Long = 66

Private Function NegZero() As Double
    Dim a As Double, b As Double
    a = -1E-300
    b = 1E-300
    NegZero = a * b                                       ' underflow to negative zero
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = Format("hello", "@@@@@@@@")               ' right-aligned, padded
        Case 2: v = Format("hello", "!@@@@@@@@")              ' left-aligned
        Case 3: v = Format("hello", "@@@")                    ' fewer placeholders than chars
        Case 4: v = Format("hello", "&&&&&&&&")
        Case 5: v = Format("hello", ">")
        Case 6: v = Format("Hello", "<")
        Case 7: v = Format("hello", ">@@@@@@@")
        Case 8: v = Format("abc", "@@@-@@@")
        Case 9: v = Format("abc", "\[@\]")
        Case 10: v = Format("", "@;empty")                    ' empty string uses second section
        Case 11: v = Format("x", "@;empty")
        Case 12: v = Format(123, "@@@@@")                     ' number under string format
        Case 13: v = Format("hello", "&")
        Case 14: v = Format("ab", "@&@")
        Case 15: v = Format("hello", "@ ""suffix""")
        Case 16: v = Format("abc", "0.00")                    ' non-numeric string, numeric format
        Case 17: v = Format("12.5", "0.00")                   ' numeric string
        Case 18: v = Format("12.5", "General Number")
        Case 19: v = Format("&H10", "0")                      ' hex string
        Case 20: v = Format(" 5 ", "0.0")
        Case 21: v = Format("1,234.5", "0.0")                 ' locale
        Case 22: v = Format("", "0.00")
        Case 23: v = Format(True, "0")
        Case 24: v = Format(True, "Yes/No")
        Case 25: v = Format(True)
        Case 26: v = Format(False, "True/False")
        Case 27: v = Format(True, "@")
        Case 28: v = Format(Empty, "0.00")
        Case 29: v = Format(Empty, "@")
        Case 30: v = Format(Empty)
        Case 31: v = Format(Null)
        Case 32: v = Format(Null, "0.00")
        Case 33: v = Format(Null, "@")
        Case 34: v = Format(CCur(1234.5678), "0.000")
        Case 35: v = Format(CCur(1234.5678))
        Case 36: v = Format(CDec("1234.565"), "0.00")
        Case 37: v = Format(CDec("79228162514264337593543950335"), "#,##0")
        Case 38: v = Format(NegZero(), "0.00")                ' negative zero
        Case 39: v = Format(NegZero())
        Case 40: v = Format(NegZero(), "0.00;(0.00);zero")
        Case 41: v = Format(1234.5)                           ' no format: like Str
        Case 42: v = Format(1234.5, "")                       ' empty format string
        Case 43: v = Format(CByte(5), "00")
        Case 44: v = Format(CInt(-5), "00;(00)")
        Case 45: v = Format(CLng(123456789), "#,##0")
        Case 46: v = Format(CSng(1.1), "0.0000000")           ' single precision
        Case 47: v = Format(CSng(1.1))
        Case 48: v = Format(1E+15)
        Case 49: v = Format(1E+14)
        Case 50: v = Format(0.1 + 0.2)
        Case 51: v = Format(123456789012345678#)
        Case 52: v = Format(1.5E-10)
        Case 53: v = Format(Array(1, 2), "0")                 ' type mismatch
        Case 54: v = Format(5, "0", 99)                       ' bad FirstDayOfWeek
        Case 55: v = Format(5, "0", , 99)                     ' bad FirstWeekOfYear
        Case 56: v = Format$(Null, "0")                       ' $ form: Null is error 94
        Case 57: v = Format$(5, "0.0")                        ' $ form
        Case 58: v = Format(5, Null)                          ' Null format
        Case 59: v = Format(CVErr(2007), "0")                 ' error value
        Case 60: v = Format(5, "0", vbMonday)                 ' firstdayofweek only
        Case 61: v = Format(5, 0)                             ' numeric format argument
        Case 62: v = Format(#1/2/2020#, "0")                  ' date under numeric format
        Case 63: v = Format(1234.5, "@")
        Case 64: v = Format(-5, "@;neg")
        Case 65: v = Format(1.5, ">")
        Case 66: v = Format(5, "General Number", vbMonday, vbFirstJan1)
    End Select
    T = "Format3." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Format3." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
