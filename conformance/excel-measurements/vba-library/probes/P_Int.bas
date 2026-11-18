Attribute VB_Name = "P_Int"
Option Explicit
' Int (VBA.Conversion; MS-VBAL 6.1.2.x).
Private Const N As Long = 38

Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, c As Currency, dt As Date, ds As Double, sg As Single
    Dim l As Long, i As Integer, by As Byte, b As Boolean, s As String, o As Object
    On Error GoTo fail
    Select Case k
        Case 1: v = Int(2.7)
        Case 2: v = Int(-2.7)
        Case 3: v = Int(2.5)
        Case 4: v = Int(-2.5)
        Case 5: v = Int(0.5)
        Case 6: v = Int(-0.5)   ' -0 or -1
        Case 7: v = Int(0)
        Case 8: v = Int(5)   ' Integer stays Integer
        Case 9: v = Int(40000)   ' Long stays Long
        Case 10: v = Int(CByte(200))   ' Byte
        Case 11: v = Int(CInt(-5))
        Case 12: v = Int(CLng(-5))
        Case 13: v = Int(CSng(-2.7))   ' Single stays Single
        Case 14: v = Int(CCur(-2.7))   ' Currency
        Case 15: v = Int(CDec("-2.7"))   ' Decimal
        Case 16: v = Int(CDec("2.999999999999999999"))
        Case 17: v = Int(#1/2/2020 3:04:05 PM#)   ' Date
        Case 18: v = Int(True)   ' Integer -1
        Case 19: v = Int(False)
        Case 20: v = Int("2.7")   ' numeric string -> Double
        Case 21: v = Int("-2.7")
        Case 22: v = Int("&HFF")   ' hex string
        Case 23: v = Int(" 3.9 ")
        Case 24: v = Int("abc")   ' 13
        Case 25: v = Int("")   ' 13
        Case 26: v = Int(Null)   ' Null
        Case 27: v = Int(Empty)   ' Integer 0
        Case 28: v = Int(CVErr(5))   ' 13
        Case 29: v = Int(1E300)
        Case 30: v = Int(-1E300)
        Case 31: v = Int(4.9E-324)
        Case 32: v = Int(123456789012.9)
        Case 33: v = Int(2147483647.9)
        Case 34: v = Int(-0#)
        Case 35: v = "compile-refused"
        Case 36: v = Int(Array(1))   ' 13
        Case 37: v = Int(0.9999999999999999)
        Case 38: v = Int(-0.9999999999999999)
    End Select
    T = "Int." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Int." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
