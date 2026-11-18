Attribute VB_Name = "P_CDbl"
Option Explicit
' CDbl (VBA.Conversion; MS-VBAL 6.1.2.x).
Private Const N As Long = 34

Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, c As Currency, dt As Date, ds As Double, sg As Single
    Dim l As Long, i As Integer, by As Byte, b As Boolean, s As String, o As Object
    On Error GoTo fail
    Select Case k
        Case 1: v = CDbl(1)
        Case 2: v = CDbl(0.1)
        Case 3: v = CDbl(CSng(0.1))   ' single widening
        Case 4: v = CDbl(-0)
        Case 5: v = CDbl(1E308)
        Case 6: v = CDbl(4.94065645841247E-324)   ' denormal min
        Case 7: v = CDbl("1.5")
        Case 8: v = CDbl(" 1.5 ")
        Case 9: v = CDbl("1.5e3")
        Case 10: v = CDbl("1.5D3")   ' D exponent
        Case 11: v = CDbl("&HFF")   ' hex
        Case 12: v = CDbl("&HFFFF")   ' hex string as Integer -1
        Case 13: v = CDbl("&O17")
        Case 14: v = CDbl("1,000")   ' locale
        Case 15: v = CDbl("abc")   ' 13
        Case 16: v = CDbl("")   ' 13
        Case 17: v = CDbl("1e309")   ' overflow 6
        Case 18: v = CDbl("1.5.2")   ' 13
        Case 19: v = CDbl("-")   ' 13
        Case 20: v = CDbl("+5")
        Case 21: v = CDbl(".5")
        Case 22: v = CDbl("5.")
        Case 23: v = CDbl("1 000")   ' 13
        Case 24: v = CDbl(True)
        Case 25: v = CDbl(Null)   ' 94
        Case 26: v = CDbl(Empty)
        Case 27: v = CDbl(CVErr(5))   ' 13
        Case 28: v = CDbl(#1/2/2020#)
        Case 29: v = CDbl(CCur(1.2345))
        Case 30: v = CDbl(CDec("1.23456789012345678"))
        Case 31: v = CDbl("TRUE")   ' 13
        Case 32: v = CDbl("1E+2")
        Case 33: v = CDbl("0.1")
        Case 34: v = CDbl("1.7976931348623157E308")
    End Select
    T = "CDbl." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "CDbl." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
