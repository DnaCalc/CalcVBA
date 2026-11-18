Attribute VB_Name = "P_CCur"
Option Explicit
' CCur (VBA.Conversion; MS-VBAL 6.1.2.x).
Private Const N As Long = 26

Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, c As Currency, dt As Date, ds As Double, sg As Single
    Dim l As Long, i As Integer, by As Byte, b As Boolean, s As String, o As Object
    On Error GoTo fail
    Select Case k
        Case 1: v = CCur(1)
        Case 2: v = CCur(1.23456)   ' 4-decimal rounding
        Case 3: v = CCur(1.00005)   ' rounding at 5th digit
        Case 4: v = CCur(1.00015)
        Case 5: v = CCur(0.00005)
        Case 6: v = CCur(-0.00005)
        Case 7: v = CCur(922337203685477.5807)   ' max
        Case 8: v = CCur(-922337203685477.5808)   ' min
        Case 9: v = CCur(922337203685478#)   ' overflow 6
        Case 10: v = CCur(1E15)   ' overflow
        Case 11: v = CCur("1.5")
        Case 12: v = CCur(" 1,000.25 ")   ' thousands separator, locale
        Case 13: v = CCur("abc")   ' 13
        Case 14: v = CCur("")   ' 13
        Case 15: v = CCur("&HFF")
        Case 16: v = CCur(True)
        Case 17: v = CCur(Null)   ' 94
        Case 18: v = CCur(Empty)
        Case 19: v = CCur(CVErr(5))   ' 13
        Case 20: v = CCur(#1/2/1900#)
        Case 21: v = CCur(CDec("1.23456789"))
        Case 22: v = CCur(0.1 + 0.2)
        Case 23: v = CCur(CSng(0.1))   ' single noise
        Case 24: v = CCur(CByte(7))
        Case 25: v = CCur("922337203685477.5807")
        Case 26: v = CCur("922337203685477.5808")   ' overflow
    End Select
    T = "CCur." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "CCur." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
