Attribute VB_Name = "P_CLngPtr"
Option Explicit
' CLngPtr (LongPtr: LongLong on 64-bit) (VBA.Conversion; MS-VBAL 6.1.2.x).
Private Const N As Long = 22

Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, c As Currency, dt As Date, ds As Double, sg As Single
    Dim l As Long, i As Integer, by As Byte, b As Boolean, s As String, o As Object
    On Error GoTo fail
    Select Case k
        Case 1: v = CLngPtr(0)
        Case 2: v = CLngPtr(2147483647)
        Case 3: v = CLngPtr(2147483648#)   ' 64-bit ok, 32-bit overflow
        Case 4: v = CLngPtr(2.5)
        Case 5: v = CLngPtr(3.5)
        Case 6: v = CLngPtr(-2.5)
        Case 7: v = CLngPtr("12")
        Case 8: v = CLngPtr("12.5")
        Case 9: v = CLngPtr("&HFFFF")
        Case 10: v = CLngPtr("&HFFFFFFFF")
        Case 11: v = CLngPtr(" 7 ")
        Case 12: v = CLngPtr("abc")   ' 13
        Case 13: v = CLngPtr("")   ' 13
        Case 14: v = CLngPtr(True)
        Case 15: v = CLngPtr(Null)   ' 94
        Case 16: v = CLngPtr(Empty)
        Case 17: v = CLngPtr(CVErr(5))   ' 13
        Case 18: v = CLngPtr(CCur(2.5))
        Case 19: v = CLngPtr(CDec(3.5))
        Case 20: v = CLngPtr(1E19)   ' overflow
        Case 21: v = CLngPtr(#1/2/2020#)
        Case 22: x = CLngPtr(5): v = TypeName(x)   ' LongLong on 64-bit
    End Select
    T = "CLngPtr." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "CLngPtr." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
