Attribute VB_Name = "P_IsNumeric"
Option Explicit
' IsNumeric (Information; numeric strings are locale dependent, marked locale).
Private Const N As Long = 55

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    Dim va As Variant, cu As Currency, dt As Date, bo As Boolean, fa(1 To 2) As Long, c As Collection, o As Object, by As Byte, sg As Single, de As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = IsNumeric(1)                              ' Integer
        Case 2: v = IsNumeric(1.5)                            ' Double
        Case 3: v = IsNumeric(CByte(1))                       ' Byte
        Case 4: v = IsNumeric(CSng(1))                        ' Single
        Case 5: cu = 1: v = IsNumeric(cu)                     ' Currency
        Case 6: de = CDec(1): v = IsNumeric(de)               ' Decimal
        Case 7: v = IsNumeric(True)                           ' Boolean is numeric
        Case 8: dt = DateSerial(2020, 1, 2): v = IsNumeric(dt)   ' Date typed value is numeric
        Case 9: v = IsNumeric("1")                            ' numeric string
        Case 10: v = IsNumeric(" 1 ")                          ' surrounding spaces
        Case 11: v = IsNumeric("1e5")                          ' exponent
        Case 12: v = IsNumeric("1E+5")
        Case 13: v = IsNumeric("1D5")                          ' D exponent
        Case 14: v = IsNumeric("&H10")                         ' hex
        Case 15: v = IsNumeric("&O17")                         ' octal
        Case 16: v = IsNumeric("&HZZ")                         ' invalid hex
        Case 17: v = IsNumeric("1,000")                        ' locale: thousands separator
        Case 18: v = IsNumeric("1.5")                          ' locale: decimal point
        Case 19: v = IsNumeric("1,5")                          ' locale: comma decimal
        Case 20: v = IsNumeric("$1")                           ' locale: currency symbol
        Case 21: v = IsNumeric("R1")                           ' locale: Rand symbol
        Case 22: v = IsNumeric("(1)")                          ' parentheses negative
        Case 23: v = IsNumeric("1-")                           ' trailing minus
        Case 24: v = IsNumeric("-1")
        Case 25: v = IsNumeric("+1")
        Case 26: v = IsNumeric("--1")
        Case 27: v = IsNumeric("1.5.2")
        Case 28: v = IsNumeric(".")
        Case 29: v = IsNumeric(".5")
        Case 30: v = IsNumeric("5.")
        Case 31: v = IsNumeric("1%")
        Case 32: v = IsNumeric("0x10")                         ' C-style hex
        Case 33: v = IsNumeric("1 2")
        Case 34: v = IsNumeric("")                             ' zero-length string
        Case 35: v = IsNumeric(" ")
        Case 36: v = IsNumeric("abc")
        Case 37: v = IsNumeric("1" & vbTab)                    ' trailing tab
        Case 38: v = IsNumeric("12345678901234567890")         ' beyond Long, within Double
        Case 39: v = IsNumeric("1E400")                        ' overflows Double
        Case 40: v = IsNumeric("Infinity")
        Case 41: v = IsNumeric("NaN")
        Case 42: v = IsNumeric("2020/01/02")                   ' locale: date string
        Case 43: v = IsNumeric("12:30")                        ' time string
        Case 44: v = IsNumeric(Null)
        Case 45: v = IsNumeric(Empty)
        Case 46: v = IsNumeric(vbNullString)
        Case 47: v = IsNumeric(Array(1, 2))                    ' array
        Case 48: v = IsNumeric(Array(1, 2)(0))                 ' array element
        Case 49: Set c = New Collection: v = IsNumeric(c)      ' object without a default property
        Case 50: v = IsNumeric(CVErr(5))                       ' error value
        Case 51: v = IsNumeric(#1/1/2000#)                     ' Date literal
        Case 52: v = IsNumeric(Chr(160) & "1")                 ' non-breaking space
        Case 53: v = IsNumeric(ChrW(&H661))                    ' Arabic-Indic digit
        Case 54: v = IsNumeric(ChrW(&HFF11))                   ' full-width digit one
        Case 55: v = IsNumeric(Expression:="1")                ' named argument
    End Select
    T = "IsNumeric." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "IsNumeric." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
