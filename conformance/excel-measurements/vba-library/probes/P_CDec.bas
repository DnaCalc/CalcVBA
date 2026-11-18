Attribute VB_Name = "P_CDec"
Option Explicit
' CDec (VBA.Conversion; MS-VBAL 6.1.2.x).
Private Const N As Long = 43

Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, c As Currency, dt As Date, ds As Double, sg As Single
    Dim l As Long, i As Integer, by As Byte, b As Boolean, s As String, o As Object
    On Error GoTo fail
    Select Case k
        Case 1: v = CDec(1)
        Case 2: v = CDec(0.1)   ' from Double
        Case 3: v = CDec(1/3)   ' Double precision 15 digits
        Case 4: v = CDec(CSng(0.1))
        Case 5: v = CDec(1E28)
        Case 6: v = CDec(7.9228162514264E+28)   ' near max
        Case 7: v = CDec(1E29)   ' overflow 6
        Case 8: v = CDec("79228162514264337593543950335")   ' max
        Case 9: v = CDec("79228162514264337593543950336")   ' overflow
        Case 10: v = CDec("-79228162514264337593543950335")
        Case 11: v = CDec("0.0000000000000000000000000001")   ' scale 28
        Case 12: v = CDec("0.00000000000000000000000000001")   ' scale 29
        Case 13: v = CDec("1.00")   ' keeps scale
        Case 14: v = CDec("1.10")
        Case 15: v = CDec("1.5")
        Case 16: v = CDec("1,5")   ' locale
        Case 17: v = CDec(" 12.5 ")
        Case 18: v = CDec("1E5")
        Case 19: v = CDec("1E-5")
        Case 20: v = CDec("&HFF")
        Case 21: v = CDec("abc")   ' 13
        Case 22: v = CDec("")   ' 13
        Case 23: v = CDec(True)
        Case 24: v = CDec(False)
        Case 25: v = CDec(Null)   ' 94
        Case 26: v = CDec(Empty)
        Case 27: v = CDec(CVErr(5))   ' 13
        Case 28: v = CDec(#1/2/2020#)
        Case 29: v = CDec(CCur(922337203685477.5807))
        Case 30: v = CDec(CCur(1.2345))
        Case 31: v = CDec(2147483647)
        Case 32: v = CDec(-0.5)
        Case 33: v = CDec(0)
        Case 34: v = CDec(-1E-28)
        Case 35: v = CDec(1E-29)   ' Double tiny
        Case 36: v = CDec(1234567890123456789#)   ' Double with 19 digits
        Case 37: v = CDec(0.1 + 0.2)
        Case 38: x = CDec("1.50"): v = x * 2   ' scale preserved in product
        Case 39: x = CDec(5): v = VarType(x)   ' vartype 14
        Case 40: v = CDec(9223372036854775807)   ' LongLong max on 64-bit
        Case 41: v = CDec(CByte(255))
        Case 42: v = CDec(" ")   ' 13
        Case 43: v = CDec("1.5E+2")
    End Select
    T = "CDec." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "CDec." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
