Attribute VB_Name = "P_CVDate"
Option Explicit
' CVDate (VBA.Conversion; MS-VBAL 6.1.2.x).
Private Const N As Long = 22

Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, c As Currency, dt As Date, ds As Double, sg As Single
    Dim l As Long, i As Integer, by As Byte, b As Boolean, s As String, o As Object
    On Error GoTo fail
    Select Case k
        Case 1: v = CVDate(0)
        Case 2: v = CVDate(1.5)
        Case 3: v = CVDate(-1.5)
        Case 4: v = CVDate(43831.75)
        Case 5: v = CVDate(2958466)   ' overflow
        Case 6: v = CVDate(#1/2/2020#)
        Case 7: v = CVDate(True)
        Case 8: v = CVDate(Null)   ' Null is returned? err 94
        Case 9: v = CVDate(Empty)
        Case 10: v = CVDate(CVErr(5))   ' 13
        Case 11: v = CVDate("")   ' 13
        Case 12: v = CVDate("abc")   ' 13
        Case 13: v = CVDate("2020/01/02")   ' locale
        Case 14: v = CVDate("12:30")
        Case 15: v = CVDate("1/2/2020")   ' locale
        Case 16: v = CVDate("43831")
        Case 17: v = CVDate("January 2, 2020")   ' locale
        Case 18: v = CVDate(CCur(1.5))
        Case 19: v = CVDate(CDec(2.25))
        Case 20: v = CVDate(CSng(1.5))
        Case 21: v = CVDate("2/29/2019")   ' 13
        Case 22: v = TypeName(CVDate(5))   ' Date
    End Select
    T = "CVDate." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "CVDate." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
