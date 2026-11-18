Attribute VB_Name = "P_CBool"
Option Explicit
' CBool (VBA.Conversion; MS-VBAL 6.1.2.x).
Private Const N As Long = 24

Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, c As Currency, dt As Date, ds As Double, sg As Single
    Dim l As Long, i As Integer, by As Byte, b As Boolean, s As String, o As Object
    On Error GoTo fail
    Select Case k
        Case 1: v = CBool(1)   ' nonzero
        Case 2: v = CBool(0)   ' zero
        Case 3: v = CBool(-1)
        Case 4: v = CBool(0.5)   ' fraction is True
        Case 5: v = CBool(0.0000001)
        Case 6: v = CBool("True")   ' string True
        Case 7: v = CBool("false")   ' string false lower case
        Case 8: v = CBool("1")   ' numeric string
        Case 9: v = CBool("0")
        Case 10: v = CBool(" 5 ")   ' spaces
        Case 11: v = CBool("abc")   ' err 13
        Case 12: v = CBool("")   ' err 13
        Case 13: v = CBool(Null)   ' err 94
        Case 14: v = CBool(Empty)
        Case 15: v = CBool(CVErr(2007))   ' Error source err 13
        Case 16: v = CBool(CCur(0))   ' currency zero
        Case 17: v = CBool(CDec(1))   ' decimal
        Case 18: v = CBool(#12/30/1899#)   ' date zero
        Case 19: v = CBool("&H0")   ' hex string
        Case 20: v = CBool("&HFF")
        Case 21: v = CBool("0.0")
        Case 22: v = CBool("1E5")
        Case 23: v = CBool(-0.0000001)
        Case 24: v = CBool("TRUE")   ' upper case
    End Select
    T = "CBool." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "CBool." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
