Attribute VB_Name = "P_CVErr"
Option Explicit
' CVErr (VBA.Conversion; MS-VBAL 6.1.2.x).
Private Const N As Long = 31

Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, c As Currency, dt As Date, ds As Double, sg As Single
    Dim l As Long, i As Integer, by As Byte, b As Boolean, s As String, o As Object
    On Error GoTo fail
    Select Case k
        Case 1: v = CVErr(0)
        Case 2: v = CVErr(1)
        Case 3: v = CVErr(5)
        Case 4: v = CVErr(7)
        Case 5: v = CVErr(2007)   ' #DIV/0!
        Case 6: v = CVErr(2042)   ' #N/A
        Case 7: v = CVErr(2015)
        Case 8: v = CVErr(65535)
        Case 9: v = CVErr(65536)   ' err 5?
        Case 10: v = CVErr(-1)   ' err 5?
        Case 11: v = CVErr(2147483647)
        Case 12: v = CVErr("5")   ' numeric string
        Case 13: v = CVErr("abc")   ' 13
        Case 14: v = CVErr("")   ' 13
        Case 15: v = CVErr(Null)   ' 94
        Case 16: v = CVErr(Empty)
        Case 17: v = CVErr(2.5)   ' rounds
        Case 18: v = CVErr(3.5)
        Case 19: v = CVErr(True)   ' -1
        Case 20: v = CVErr(CByte(5))
        Case 21: v = CVErr(#1/2/1900#)
        Case 22: v = IsError(CVErr(5))
        Case 23: v = TypeName(CVErr(5))
        Case 24: v = VarType(CVErr(2042))
        Case 25: x = CVErr(2042): v = CLng(x)   ' Error to Long
        Case 26: x = CVErr(2042): v = CStr(x)   ' Error to String
        Case 27: x = CVErr(7): v = x & ""   ' concatenation raises
        Case 28: x = CVErr(5): v = x = x   ' comparison raises
        Case 29: x = CVErr(5): v = IsNull(x)
        Case 30: v = CVErr(CVErr(5))   ' nested
        Case 31: v = CVErr(CCur(7.5))
    End Select
    T = "CVErr." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "CVErr." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
