Attribute VB_Name = "P_Fix"
Option Explicit
' Fix (VBA.Conversion; MS-VBAL 6.1.2.x).
Private Const N As Long = 38

Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, c As Currency, dt As Date, ds As Double, sg As Single
    Dim l As Long, i As Integer, by As Byte, b As Boolean, s As String, o As Object
    On Error GoTo fail
    Select Case k
        Case 1: v = Fix(2.7)
        Case 2: v = Fix(-2.7)
        Case 3: v = Fix(2.5)
        Case 4: v = Fix(-2.5)
        Case 5: v = Fix(0.5)
        Case 6: v = Fix(-0.5)   ' -0 or -1
        Case 7: v = Fix(0)
        Case 8: v = Fix(5)   ' Integer stays Integer
        Case 9: v = Fix(40000)   ' Long stays Long
        Case 10: v = Fix(CByte(200))   ' Byte
        Case 11: v = Fix(CInt(-5))
        Case 12: v = Fix(CLng(-5))
        Case 13: v = Fix(CSng(-2.7))   ' Single stays Single
        Case 14: v = Fix(CCur(-2.7))   ' Currency
        Case 15: v = Fix(CDec("-2.7"))   ' Decimal
        Case 16: v = Fix(CDec("2.999999999999999999"))
        Case 17: v = Fix(#1/2/2020 3:04:05 PM#)   ' Date
        Case 18: v = Fix(True)   ' Integer -1
        Case 19: v = Fix(False)
        Case 20: v = Fix("2.7")   ' numeric string -> Double
        Case 21: v = Fix("-2.7")
        Case 22: v = Fix("&HFF")   ' hex string
        Case 23: v = Fix(" 3.9 ")
        Case 24: v = Fix("abc")   ' 13
        Case 25: v = Fix("")   ' 13
        Case 26: v = Fix(Null)   ' Null
        Case 27: v = Fix(Empty)   ' Integer 0
        Case 28: v = Fix(CVErr(5))   ' 13
        Case 29: v = Fix(1E300)
        Case 30: v = Fix(-1E300)
        Case 31: v = Fix(4.9E-324)
        Case 32: v = Fix(123456789012.9)
        Case 33: v = Fix(2147483647.9)
        Case 34: v = Fix(-0#)
        Case 35: v = "compile-refused"
        Case 36: v = Fix(Array(1))   ' 13
        Case 37: v = Fix(0.9999999999999999)
        Case 38: v = Fix(-0.9999999999999999)
    End Select
    T = "Fix." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Fix." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
