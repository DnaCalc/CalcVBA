Attribute VB_Name = "P_CStr"
Option Explicit
' CStr (VBA.Conversion; MS-VBAL 6.1.2.x).
Private Const N As Long = 72

Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, c As Currency, dt As Date, ds As Double, sg As Single
    Dim l As Long, i As Integer, by As Byte, b As Boolean, s As String, o As Object
    On Error GoTo fail
    Select Case k
        Case 1: v = CStr(0)
        Case 2: v = CStr(1)
        Case 3: v = CStr(-1)
        Case 4: v = CStr(1.5)   ' locale decimal
        Case 5: v = CStr(0.1)
        Case 6: v = CStr(1/3)   ' 15 digits
        Case 7: v = CStr(0.1 + 0.2)
        Case 8: v = CStr(100000)
        Case 9: v = CStr(1E15)
        Case 10: v = CStr(1E16)
        Case 11: v = CStr(123456789012345#)
        Case 12: v = CStr(1234567890123456#)
        Case 13: v = CStr(12345678901234567#)
        Case 14: v = CStr(1E-5)
        Case 15: v = CStr(0.0001)
        Case 16: v = CStr(0.00001234)
        Case 17: v = CStr(1.5E+300)
        Case 18: v = CStr(-1.5E-300)
        Case 19: v = CStr(1.7976931348623157E+308)
        Case 20: v = CStr(4.94065645841247E-324)
        Case 21: v = CStr(123456.789)
        Case 22: v = CStr(1E+14)
        Case 23: v = CStr(1.5E+15)
        Case 24: v = CStr(999999999999999#)
        Case 25: v = CStr(0.99999999999999989)
        Case 26: v = CStr(2147483647)
        Case 27: v = CStr(CLng(-2147483648#))
        Case 28: v = CStr(CInt(5))
        Case 29: v = CStr(CByte(5))
        Case 30: v = CStr(CSng(0.1))   ' single 7 digits
        Case 31: v = CStr(CSng(1/3))
        Case 32: v = CStr(CSng(1E10))
        Case 33: v = CStr(CSng(1.5E-10))
        Case 34: v = CStr(CSng(16777216))
        Case 35: v = CStr(CSng(123456.7))
        Case 36: v = CStr(CSng(3.40282347E+38))
        Case 37: v = CStr(CCur(1.5))
        Case 38: v = CStr(CCur(0.0001))
        Case 39: v = CStr(CCur(-922337203685477.5807))
        Case 40: v = CStr(CCur(1234567.891))
        Case 41: v = CStr(CDec("1.50"))   ' scale kept
        Case 42: v = CStr(CDec("79228162514264337593543950335"))
        Case 43: v = CStr(CDec("0.0000000000000000000000000001"))
        Case 44: v = CStr(CDec(0))
        Case 45: v = CStr(CDec("-1.5"))
        Case 46: v = CStr(True)   ' locale
        Case 47: v = CStr(False)
        Case 48: v = CStr(Null)   ' err 94
        Case 49: v = CStr(Empty)   ' empty string
        Case 50: v = CStr(CVErr(2007))   ' Error 2007
        Case 51: v = CStr(CVErr(5))
        Case 52: v = CStr("abc")
        Case 53: v = CStr("")
        Case 54: v = CStr(#1/2/2020#)   ' locale short date
        Case 55: v = CStr(#1/2/2020 3:04:05 PM#)   ' locale
        Case 56: v = CStr(#3:04:05 PM#)   ' time only, locale
        Case 57: v = CStr(#12/30/1899#)   ' zero date, locale
        Case 58: v = CStr(CDate(0.5))   ' locale
        Case 59: v = CStr(CDate(1.0000001))   ' locale
        Case 60: v = CStr(CDate(-1.5))   ' locale
        Case 61: v = CStr(#12/31/9999 11:59:59 PM#)   ' locale
        Case 62: v = CStr(#1/1/0100#)   ' locale
        Case 63: v = CStr(Array(1))   ' err 13
        Case 64: Set o = New Collection: v = CStr(o)   ' Collection has no default value: err 438 or 450
        Case 65: x = 5: v = CStr(x)   ' Variant Integer
        Case 66: x = 1.5@: v = CStr(x)   ' Variant Currency
        Case 67: v = CStr(-0.5)
        Case 68: v = CStr(5E-324 * 0)
        Case 69: v = CStr(-0)
        Case 70: v = CStr(1.5@)
        Case 71: v = CStr(123456789.123456789)
        Case 72: v = CStr(0.000123456789012345678)
    End Select
    T = "CStr." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "CStr." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
