Attribute VB_Name = "P_CLng"
Option Explicit
' CLng (VBA.Conversion; MS-VBAL 6.1.2.x).
Private Const N As Long = 34

Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, c As Currency, dt As Date, ds As Double, sg As Single
    Dim l As Long, i As Integer, by As Byte, b As Boolean, s As String, o As Object
    On Error GoTo fail
    Select Case k
        Case 1: v = CLng(0)
        Case 2: v = CLng(2147483647)
        Case 3: v = CLng(-2147483648#)
        Case 4: v = CLng(2147483648#)   ' overflow
        Case 5: v = CLng(-2147483649#)   ' overflow
        Case 6: v = CLng(2.5)
        Case 7: v = CLng(3.5)
        Case 8: v = CLng(-2.5)
        Case 9: v = CLng(-3.5)
        Case 10: v = CLng(0.5)
        Case 11: v = CLng(1.5)
        Case 12: v = CLng(2147483647.5)   ' rounds to 2^31 overflow
        Case 13: v = CLng(2147483646.5)   ' banker even 2147483646
        Case 14: v = CLng(-2147483648.5)   ' banker to -2147483648
        Case 15: v = CLng("12")
        Case 16: v = CLng(" 12 ")
        Case 17: v = CLng("12.5")
        Case 18: v = CLng("&HFFFFFFFF")   ' hex string -1
        Case 19: v = CLng("&HFFFF")   ' 65535
        Case 20: v = CLng("&H7FFFFFFF")
        Case 21: v = CLng("&H100000000")   ' overflow
        Case 22: v = CLng("&O777")
        Case 23: v = CLng("1E3")
        Case 24: v = CLng("abc")   ' 13
        Case 25: v = CLng("")   ' 13
        Case 26: v = CLng(True)
        Case 27: v = CLng(Null)   ' 94
        Case 28: v = CLng(Empty)
        Case 29: v = CLng(CVErr(5))   ' 13
        Case 30: v = CLng(CCur(2.5))
        Case 31: v = CLng(CDec(3.5))
        Case 32: v = CLng(CSng(2.5))
        Case 33: v = CLng(#1/2/2020#)
        Case 34: v = CLng(CSng(16777217))   ' single precision
    End Select
    T = "CLng." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "CLng." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
