Attribute VB_Name = "P_CVar"
Option Explicit
' CVar (VBA.Conversion; MS-VBAL 6.1.2.x).
Private Const N As Long = 23

Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, c As Currency, dt As Date, ds As Double, sg As Single
    Dim l As Long, i As Integer, by As Byte, b As Boolean, s As String, o As Object
    On Error GoTo fail
    Select Case k
        Case 1: v = CVar(1)
        Case 2: v = CVar(1.5)
        Case 3: v = CVar("a")
        Case 4: v = CVar("")
        Case 5: v = CVar(Null)
        Case 6: v = CVar(Empty)
        Case 7: v = CVar(True)
        Case 8: v = CVar(CVErr(5))
        Case 9: v = CVar(#1/2/2020#)
        Case 10: v = CVar(CCur(1.5))
        Case 11: v = CVar(CDec(2.5))
        Case 12: v = CVar(CSng(0.1))
        Case 13: v = CVar(CByte(5))
        Case 14: v = CVar(CInt(5))
        Case 15: v = CVar(CLng(5))
        Case 16: v = CVar(2147483648#)
        Case 17: v = CVar(Array(1, 2))
        Case 18: v = CVar(-0)
        Case 19: v = CVar(12345678901)
        Case 20: x = 5: v = CVar(x)   ' already Variant
        Case 21: Set v = CVar(New Collection)   ' object
        Case 22: v = CVar(CVar(CByte(9)))
        Case 23: v = CVar(CLngPtr(5))
    End Select
    T = "CVar." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "CVar." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
