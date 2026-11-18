Attribute VB_Name = "P_CCurEdges"
Option Explicit
' Excel-backed Currency binary rounding, exact text, and assignment neighbors.
Private Const N As Long = 110
Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, c As Currency
    On Error GoTo fail
    Select Case k
        Case 1: v = CCur(0.00005)
        Case 2: c = 0.00005: v = c
        Case 3: v = CCur(-0.00005)
        Case 4: c = -0.00005: v = c
        Case 5: v = CCur(0.00015)
        Case 6: c = 0.00015: v = c
        Case 7: v = CCur(-0.00015)
        Case 8: c = -0.00015: v = c
        Case 9: v = CCur(1.00005)
        Case 10: c = 1.00005: v = c
        Case 11: v = CCur(1.00015)
        Case 12: c = 1.00015: v = c
        Case 13: v = CCur(0.00025)
        Case 14: c = 0.00025: v = c
        Case 15: v = CCur(0.00035)
        Case 16: c = 0.00035: v = c
        Case 17: v = CCur(0.00125)
        Case 18: c = 0.00125: v = c
        Case 19: v = CCur(0.00135)
        Case 20: c = 0.00135: v = c
        Case 21: v = CCur(922337203685477.5#)
        Case 22: c = 922337203685477.5#: v = c
        Case 23: v = CCur(-922337203685477.5#)
        Case 24: c = -922337203685477.5#: v = c
        Case 25: v = CCur(922337203685477.625#)
        Case 26: c = 922337203685477.625#: v = c
        Case 27: v = CCur(-922337203685477.625#)
        Case 28: c = -922337203685477.625#: v = c
        Case 29: v = CCur(1E-300)
        Case 30: c = 1E-300: v = c
        Case 31: v = CCur("0.00005")
        Case 32: v = CCur("-0.00005")
        Case 33: v = CCur("0.00015")
        Case 34: v = CCur("-0.00015")
        Case 35: v = CCur("1.00005")
        Case 36: v = CCur("1.00015")
        Case 37: v = CCur("922337203685477.5807")
        Case 38: v = CCur("922337203685477.5808")
        Case 39: v = CCur("-922337203685477.5808")
        Case 40: v = CCur("-922337203685477.5809")
        Case 41: v = CCur("922337203685477.58065")
        Case 42: v = CCur("922337203685477.58075")
        Case 43: v = CCur("-922337203685477.58085")
        Case 44: v = CCur("0.0000500000000000000000000000000001")
        Case 45: v = CCur("0.0001499999999999999999999999999999")
        Case 46: v = CCur("1E-5")
        Case 47: v = CCur("1E-4")
        Case 48: v = CCur("1E15")
        Case 49: v = CCur("1E300")
        Case 50: v = CCur("1E-300")
        Case 51: v = CCur("&HFFFF")
        Case 52: v = CCur("&HFFFFFFFF")
        Case 53: v = CCur(" 1,000.25 ")
        Case 54: v = CCur("1.2.3")
        Case 55: v = CCur(CDec("0.00005"))
        Case 56: v = CCur(CDec("0.00015"))
        Case 57: v = CCur(CSng(0.00005))
        Case 58: v = CCur(CLngLng("922337203685477"))
        Case 59: v = CCur(CVErr(5))
        Case 60: x = "0.00005": c = x: v = c
        Case 61: x = "0.00015": c = x: v = c
        Case 62: x = "922337203685477.5807": c = x: v = c
        Case 63: x = "-922337203685477.5808": c = x: v = c
        Case 64: x = "922337203685477.5808": c = x: v = c
        Case 65: x = "0.0000500000000000000000000000000001": c = x: v = c
        Case 66: x = "0.0001499999999999999999999999999999": c = x: v = c
        Case 67: x = "&HFFFF": c = x: v = c
        Case 68: x = "1E-300": c = x: v = c
        Case 69: v = CCur("(1.00005)")
        Case 70: v = CCur("1.00015-")
        Case 71: v = CCur("- 0.00005")
        Case 72: v = CCur("1 000.00005")
        Case 73: v = CCur("1" & ChrW(160) & "000.00005")
        Case 74: v = CCur("1D-4")
        Case 75: v = CCur("0E9999999999999999999999")
        Case 76: v = CCur("1E9999999999999999999999")
        Case 77: v = CCur("1E-9999999999999999999999")
        Case 78: v = CCur("922337203685477.580749999999999999")
        Case 79: v = CCur("-922337203685477.58085000000000001")
        Case 80: v = CCur("1e309")
        Case 81: v = CCur("1e999")
        Case 82: v = CCur("1e1000")
        Case 83: v = CCur("1e10000")
        Case 84: v = CCur("1e32767")
        Case 85: v = CCur("1e32768")
        Case 86: v = CCur("1e2147483647")
        Case 87: v = CCur("1e2147483648")
        Case 88: v = CCur("1e-1000")
        Case 89: v = CCur("1e-32768")
        Case 90: v = CCur("1e-2147483648")
        Case 91: v = CCur("0e10000")
        Case 92: v = CCur("1e0000000000000000000000001")
        Case 93: v = CCur("1e-999")
        Case 94: v = CCur("1e-1001")
        Case 95: v = CCur("1e1073741819")
        Case 96: v = CCur("1e1073741820")
        Case 97: v = CCur("1e-1073741819")
        Case 98: v = CCur("1e-1073741820")
        Case 99: v = CCur("0e1073741819")
        Case 100: v = CCur("0e1073741820")
        Case 101: v = CCur("0e14")
        Case 102: v = CCur("0e15")
        Case 103: v = CCur("0.0e15")
        Case 104: v = CCur("000e14")
        Case 105: v = CCur("000e15")
        Case 106: v = CCur("0.0000e15")
        Case 107: v = CCur("0.0000e19")
        Case 108: v = CCur("00.00e15")
        Case 109: v = CCur("000.000e15")
        Case 110: v = CCur("0e00000000000000000000000014")
    End Select
    T = "CCurEdges." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "CCurEdges." & k & "=ERR" & Err.Number
End Function
Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
