Attribute VB_Name = "P_Val"
Option Explicit
' Val (VBA.Conversion; MS-VBAL 6.1.2.x).
Private Const N As Long = 95

Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, c As Currency, dt As Date, ds As Double, sg As Single
    Dim l As Long, i As Integer, by As Byte, b As Boolean, s As String, o As Object
    On Error GoTo fail
    Select Case k
        Case 1: v = Val("12")
        Case 2: v = Val("-12")
        Case 3: v = Val("+12")
        Case 4: v = Val("  12")   ' leading spaces
        Case 5: v = Val("12  ")
        Case 6: v = Val("1 2 3")   ' spaces inside skipped
        Case 7: v = Val("12abc")   ' stops at letter
        Case 8: v = Val("abc")   ' 0
        Case 9: v = Val("")   ' 0
        Case 10: v = Val(" ")   ' 0
        Case 11: v = Val("1.5")
        Case 12: v = Val(".5")
        Case 13: v = Val("5.")
        Case 14: v = Val("-.5")
        Case 15: v = Val("1.2.3")   ' second decimal point stops
        Case 16: v = Val("1,000")   ' comma stops
        Case 17: v = Val("1e3")
        Case 18: v = Val("1E3")
        Case 19: v = Val("1d3")   ' D exponent
        Case 20: v = Val("1D3")
        Case 21: v = Val("1e+3")
        Case 22: v = Val("1e-3")
        Case 23: v = Val("1d-3")
        Case 24: v = Val("1e")
        Case 25: v = Val("1e+")
        Case 26: v = Val("1ex")
        Case 27: v = Val("1e400")   ' overflow 6
        Case 28: v = Val("1e-400")   ' underflow
        Case 29: v = Val("1.5e3.5")
        Case 30: v = Val("&HFF")   ' hex
        Case 31: v = Val("&hff")
        Case 32: v = Val("&HFFFF")   ' Integer -1
        Case 33: v = Val("&H8000")
        Case 34: v = Val("&H10000")   ' Long 65536
        Case 35: v = Val("&HFFFFFFFF")   ' Long -1
        Case 36: v = Val("&H7FFFFFFF")
        Case 37: v = Val("&H100000000")   ' overflow?
        Case 38: v = Val("&H")   ' 0
        Case 39: v = Val("&HG")   ' 0
        Case 40: v = Val("&HFFZ")   ' stops at Z
        Case 41: v = Val("&H F F")   ' spaces inside hex
        Case 42: v = Val("&O17")   ' octal
        Case 43: v = Val("&o17")
        Case 44: v = Val("&O177777")   ' Integer -1
        Case 45: v = Val("&O8")   ' 0
        Case 46: v = Val("&O")   ' 0
        Case 47: v = Val("&B101")   ' not binary: 0
        Case 48: v = Val("&17")   ' 0
        Case 49: v = Val("12%")   ' type suffix
        Case 50: v = Val("12&")
        Case 51: v = Val("12!")
        Case 52: v = Val("12#")
        Case 53: v = Val("12@")
        Case 54: v = Val("12$")   ' $ not numeric
        Case 55: v = Val("12%3")   ' stops after suffix?
        Case 56: v = Val("&HFF%")
        Case 57: v = Val("&HFF&")
        Case 58: v = Val("&HFF!")
        Case 59: v = Val("- 5")   ' space after sign
        Case 60: v = Val("--5")
        Case 61: v = Val("+-5")
        Case 62: v = Val("5-")
        Case 63: v = Val("0x10")   ' 0
        Case 64: v = Val("00012")
        Case 65: v = Val("1" & vbTab & "2")   ' tab skipped?
        Case 66: v = Val("1" & vbCr & vbLf & "2")   ' newline skipped?
        Case 67: v = Val("1" & ChrW(160) & "2")   ' nbsp
        Case 68: v = Val("1" & ChrW(&H3000) & "2")   ' ideographic space
        Case 69: v = Val("123456789012345678901234567890")
        Case 70: v = Val("0.1")
        Case 71: v = Val("0.30000000000000004")
        Case 72: v = Val("4.9E-324")
        Case 73: v = Val("1.7976931348623157E308")
        Case 74: v = Val("9223372036854775807")
        Case 75: v = Val("  &H10  ")   ' spaces around hex
        Case 76: v = Val("1e5e5")
        Case 77: v = Val(".")   ' 0
        Case 78: v = Val("e5")   ' 0
        Case 79: v = Val("-")   ' 0
        Case 80: v = Val("1.")   ' 1
        Case 81: v = Val("1.e2")
        Case 82: v = Val(".e2")
        Case 83: v = Val(12.5)   ' number to string
        Case 84: v = Val(True)   ' "True" -> 0
        Case 85: v = Val(Null)   ' 94
        Case 86: v = Val(Empty)   ' 0 or 13
        Case 87: v = Val(CVErr(5))   ' 13
        Case 88: v = Val(#1/2/2020#)   ' date string, locale
        Case 89: v = Val(String:="12")   ' named argument
        Case 90: v = Val(Array(1))   ' 13
        Case 91: v = Val(CByte(5))
        Case 92: v = Val(CCur(1.5))
        Case 93: v = Val(CDec("1.5"))
        Case 94: v = Val(1E20)   ' Double to String "1E+20"
        Case 95: v = Val(-0.000001)   ' "-1E-06"
    End Select
    T = "Val." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Val." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
