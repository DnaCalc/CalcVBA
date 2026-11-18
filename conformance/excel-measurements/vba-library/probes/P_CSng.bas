Attribute VB_Name = "P_CSng"
Option Explicit
' CSng (VBA.Conversion; MS-VBAL 6.1.2.x).
Private Const N As Long = 33

Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, c As Currency, dt As Date, ds As Double, sg As Single
    Dim l As Long, i As Integer, by As Byte, b As Boolean, s As String, o As Object
    On Error GoTo fail
    Select Case k
        Case 1: v = CSng(1)
        Case 2: v = CSng(0.1)   ' rounded to single
        Case 3: v = CSng(1/3)
        Case 4: v = CSng(16777217)   ' 2^24+1
        Case 5: v = CSng(3.40282347E+38)   ' max
        Case 6: v = CSng(3.5E+38)   ' overflow 6
        Case 7: v = CSng(-3.5E+38)   ' overflow
        Case 8: v = CSng(1E-45)   ' denormal min
        Case 9: v = CSng(1E-46)   ' underflow to 0
        Case 10: v = CSng(1.401298E-45)
        Case 11: v = CSng(3.40282356779733E+38)   ' rounds to max or overflows
        Case 12: v = CSng(CDbl(0.1))
        Case 13: v = CSng(-0)
        Case 14: v = CSng("1.5")
        Case 15: v = CSng(" 1.5 ")
        Case 16: v = CSng("0.1")
        Case 17: v = CSng("1.5E3")
        Case 18: v = CSng("1.5D3")
        Case 19: v = CSng("&HFF")
        Case 20: v = CSng("abc")   ' 13
        Case 21: v = CSng("")   ' 13
        Case 22: v = CSng("3.5E38")   ' overflow
        Case 23: v = CSng(True)
        Case 24: v = CSng(Null)   ' 94
        Case 25: v = CSng(Empty)
        Case 26: v = CSng(CVErr(5))   ' 13
        Case 27: v = CSng(CCur(1.2345))
        Case 28: v = CSng(CDec("1.23456789"))
        Case 29: v = CSng(#1/2/2020#)
        Case 30: v = CSng(123456789)
        Case 31: v = CSng(2147483647)
        Case 32: v = CSng(8388608.5)   ' tie at 2^23
        Case 33: v = CSng(16777219)   ' tie rounds to even
    End Select
    T = "CSng." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "CSng." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
