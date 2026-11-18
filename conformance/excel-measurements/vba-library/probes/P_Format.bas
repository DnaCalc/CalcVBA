Attribute VB_Name = "P_Format"
Option Explicit
' Format / Format$ part 1: numeric formats (MS-VBAL 6.1.2.11.1.x).
Private Const N As Long = 66

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = Format(1234.5678, "General Number")       ' named
        Case 2: v = Format(1234.5678, "Currency")             ' locale
        Case 3: v = Format(1234.5678, "Fixed")
        Case 4: v = Format(1234.5678, "Standard")
        Case 5: v = Format(0.256, "Percent")
        Case 6: v = Format(1234.5678, "Scientific")
        Case 7: v = Format(1, "Yes/No")
        Case 8: v = Format(0, "Yes/No")
        Case 9: v = Format(5, "True/False")
        Case 10: v = Format(0, "On/Off")
        Case 11: v = Format(-0.00001234, "General Number")    ' small number
        Case 12: v = Format(-1234.5, "Currency")              ' locale: negative currency
        Case 13: v = Format(3.14159, "0.00")
        Case 14: v = Format(42, "00000")                      ' zero padding
        Case 15: v = Format(0.5, "#.##")
        Case 16: v = Format(1234567.891, "#,##0")             ' thousands separator
        Case 17: v = Format(1234567, "#,##0,")                ' trailing comma scales by 1000
        Case 18: v = Format(1234567890, "0.0,,")              ' two trailing commas
        Case 19: v = Format(0.256, "0%")
        Case 20: v = Format(12345.678, "0.00E+00")
        Case 21: v = Format(0.00012345, "0.00e-00")
        Case 22: v = Format(12345, "0.00E-00")
        Case 23: v = Format(12345, "0.00e+00")
        Case 24: v = Format(1234.5, "$#,##0.00")              ' literal $
        Case 25: v = Format(5, "0 \u\n\i\t\s")                ' backslash escapes
        Case 26: v = Format(5, "0 ""units""")                 ' quoted text
        Case 27: v = Format(-5.5, "#,##0.00;(#,##0.00)")      ' two sections
        Case 28: v = Format(0, "0;-0;""zero""")               ' zero section
        Case 29: v = Format(-5, "0.00;;")                     ' empty negative section
        Case 30: v = Format(Null, "0;-0;Zero;Nothing")        ' fourth section for Null
        Case 31: v = Format(0.5, "0")                         ' rounding at .5
        Case 32: v = Format(1.5, "0")
        Case 33: v = Format(2.5, "0")
        Case 34: v = Format(0.125, "0.00")
        Case 35: v = Format(2.675, "0.00")
        Case 36: v = Format(-2.5, "0")
        Case 37: v = Format(1E+300, "General Number")         ' very large
        Case 38: v = Format(1E-300, "0.00")                   ' very small
        Case 39: v = Format(1E+15, "#,##0")
        Case 40: v = Format(123456789012345678#, "0")
        Case 41: v = Format(0.5, "#.00")
        Case 42: v = Format(5, "0.#")                         ' trailing decimal point
        Case 43: v = Format(0, "#")                           ' empty result
        Case 44: v = Format(0, "#,###")
        Case 45: v = Format(5551234567#, "(000) 000-0000")
        Case 46: v = Format(-0.001, "0.00")                   ' negative that rounds to zero
        Case 47: v = Format(1234.5, "#,##0.00;(#,##0.00);""nil""")
        Case 48: v = Format(0.000123, "0.00E+00")
        Case 49: v = Format(1234.5678, "00.0#")
        Case 50: v = Format(12, "#0.0#")
        Case 51: v = Format(99.999, "0.00")                   ' carry
        Case 52: v = Format(0.999, "0%")
        Case 53: v = Format(1234.5, "0.00 ""USD""")
        Case 54: v = Format(1234.5, "\#0.00")                 ' escaped hash
        Case 55: v = Format(1234.5, "#,##0.00-")
        Case 56: v = Format(1234.5, "0.0,")                   ' comma after decimal scales
        Case 57: v = Format(1234.5, "0,0")                    ' embedded comma
        Case 58: v = Format(-1234.5, "0.00;-0.00;0.00")
        Case 59: v = Format(1234.5, "General Number;x")       ' named + extra
        Case 60: v = Format(-1, "0;(0)")
        Case 61: v = Format$(1234.5678, "0.000")              ' $ form: String
        Case 62: v = Format(Expression:=3.5, Format:="0.0")   ' named arguments
        Case 63: v = Format(12345.6789, "0.0000E+0")
        Case 64: v = Format(0.1 + 0.2, "0.000000000000000000")
        Case 65: v = Format(1234.5, "% 0.0")                  ' percent anywhere
        Case 66: v = Format(1, "#.#%")
    End Select
    T = "Format." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Format." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
