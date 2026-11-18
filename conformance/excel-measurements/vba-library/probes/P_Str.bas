Attribute VB_Name = "P_Str"
Option Explicit
' Str / Str$ (VBA.Conversion; MS-VBAL 6.1.2.x).
Private Const N As Long = 49

Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, c As Currency, dt As Date, ds As Double, sg As Single
    Dim l As Long, i As Integer, by As Byte, b As Boolean, s As String, o As Object
    On Error GoTo fail
    Select Case k
        Case 1: v = Str(5)   ' leading space
        Case 2: v = Str(-5)
        Case 3: v = Str(0)   ' leading space
        Case 4: v = Str(1.5)
        Case 5: v = Str(-1.5)
        Case 6: v = Str(0.1)
        Case 7: v = Str(1/3)
        Case 8: v = Str(1E20)
        Case 9: v = Str(1E-10)
        Case 10: v = Str(123456789012345#)
        Case 11: v = Str(1.7976931348623157E+308)
        Case 12: v = Str(100000)
        Case 13: v = Str(2147483648#)
        Case 14: v = Str(CByte(5))
        Case 15: v = Str(CInt(-5))
        Case 16: v = Str(CLng(5))
        Case 17: v = Str(CSng(0.1))
        Case 18: v = Str(CSng(1E10))
        Case 19: v = Str(CCur(1.5))
        Case 20: v = Str(CCur(-1.5))
        Case 21: v = Str(CDec("1.50"))
        Case 22: v = Str(CDec("-0.0000000000000000000000000001"))
        Case 23: v = Str(True)   ' Boolean
        Case 24: v = Str(False)
        Case 25: v = Str(#1/2/2020#)   ' locale
        Case 26: v = Str(#3:04:05 PM#)   ' locale
        Case 27: v = Str(Null)   ' Null
        Case 28: v = Str(Empty)   ' " 0"
        Case 29: v = Str("12")   ' numeric string
        Case 30: v = Str("-12")
        Case 31: v = Str(" 12 ")
        Case 32: v = Str("abc")   ' 13
        Case 33: v = Str("")   ' 13
        Case 34: v = Str("&HFF")   ' hex string
        Case 35: v = Str("1.5E3")
        Case 36: v = Str(CVErr(5))   ' 13
        Case 37: v = Str(Number:=5)   ' named argument
        Case 38: v = Str$(5)   ' String form
        Case 39: v = Str$(-5)
        Case 40: v = Str$(1.5)
        Case 41: v = Str$(0)
        Case 42: v = Str$(Null)   ' 94
        Case 43: v = Str$(Number:=7)
        Case 44: v = Str$(True)
        Case 45: v = Str(Array(1))   ' 13
        Case 46: v = Str(0.00001)
        Case 47: v = Str(-0)
        Case 48: v = Str(1E15)
        Case 49: v = Str(1E14)
    End Select
    T = "Str." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Str." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
