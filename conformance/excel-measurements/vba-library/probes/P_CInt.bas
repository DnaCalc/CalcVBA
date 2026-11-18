Attribute VB_Name = "P_CInt"
Option Explicit
' CInt (VBA.Conversion; MS-VBAL 6.1.2.x).
Private Const N As Long = 37

Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, c As Currency, dt As Date, ds As Double, sg As Single
    Dim l As Long, i As Integer, by As Byte, b As Boolean, s As String, o As Object
    On Error GoTo fail
    Select Case k
        Case 1: v = CInt(0)
        Case 2: v = CInt(32767)
        Case 3: v = CInt(-32768)
        Case 4: v = CInt(32768)   ' overflow
        Case 5: v = CInt(-32769)   ' overflow
        Case 6: v = CInt(2.5)   ' banker 2
        Case 7: v = CInt(3.5)   ' 4
        Case 8: v = CInt(-2.5)   ' -2
        Case 9: v = CInt(-3.5)   ' -4
        Case 10: v = CInt(0.5)
        Case 11: v = CInt(1.5)
        Case 12: v = CInt(32767.4)
        Case 13: v = CInt(32767.5)   ' rounds to 32768 overflow
        Case 14: v = CInt(-32768.5)   ' rounds to -32768
        Case 15: v = CInt(2.50000001)
        Case 16: v = CInt(2.4999999)
        Case 17: v = CInt("12")
        Case 18: v = CInt(" 12 ")
        Case 19: v = CInt("12.5")   ' string to double then banker
        Case 20: v = CInt("&HFF")
        Case 21: v = CInt("&HFFFF")   ' hex string -1
        Case 22: v = CInt("&H8000")   ' -32768
        Case 23: v = CInt("&H10000")   ' overflow
        Case 24: v = CInt("&O17")
        Case 25: v = CInt("1E2")
        Case 26: v = CInt("abc")   ' 13
        Case 27: v = CInt("")   ' 13
        Case 28: v = CInt(True)
        Case 29: v = CInt(Null)   ' 94
        Case 30: v = CInt(Empty)
        Case 31: v = CInt(CVErr(5))   ' 13
        Case 32: v = CInt(CCur(2.5))
        Case 33: v = CInt(CCur(3.5))
        Case 34: v = CInt(CDec(2.5))
        Case 35: v = CInt(CDec("2.5000001"))
        Case 36: v = CInt(CSng(2.5))
        Case 37: v = CInt(#1/2/1900#)
    End Select
    T = "CInt." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "CInt." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
