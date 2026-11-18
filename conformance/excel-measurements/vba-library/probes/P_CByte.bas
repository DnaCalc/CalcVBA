Attribute VB_Name = "P_CByte"
Option Explicit
' CByte (VBA.Conversion; MS-VBAL 6.1.2.x).
Private Const N As Long = 28

Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, c As Currency, dt As Date, ds As Double, sg As Single
    Dim l As Long, i As Integer, by As Byte, b As Boolean, s As String, o As Object
    On Error GoTo fail
    Select Case k
        Case 1: v = CByte(0)
        Case 2: v = CByte(255)   ' max
        Case 3: v = CByte(256)   ' overflow 6
        Case 4: v = CByte(-1)   ' overflow 6
        Case 5: v = CByte(2.5)   ' banker 2
        Case 6: v = CByte(3.5)   ' banker 4
        Case 7: v = CByte(0.5)   ' banker 0
        Case 8: v = CByte(1.5)
        Case 9: v = CByte(254.5)   ' banker 254
        Case 10: v = CByte(255.5)   ' rounds to 256: overflow
        Case 11: v = CByte(-0.5)   ' rounds to 0
        Case 12: v = CByte(-0.6)   ' rounds to -1 overflow
        Case 13: v = CByte("12")
        Case 14: v = CByte(" 12 ")
        Case 15: v = CByte("&HFF")   ' hex string
        Case 16: v = CByte("&O17")   ' octal string
        Case 17: v = CByte("1e2")
        Case 18: v = CByte("abc")   ' err 13
        Case 19: v = CByte("")   ' err 13
        Case 20: v = CByte(True)   ' -1 overflow
        Case 21: v = CByte(False)
        Case 22: v = CByte(Null)   ' err 94
        Case 23: v = CByte(Empty)
        Case 24: v = CByte(CVErr(5))   ' err 13
        Case 25: v = CByte(CCur(7.5))
        Case 26: v = CByte(CDec(8.5))
        Case 27: v = CByte(#1/2/1900#)   ' date serial 3
        Case 28: v = CByte(CSng(2.5))
    End Select
    T = "CByte." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "CByte." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
