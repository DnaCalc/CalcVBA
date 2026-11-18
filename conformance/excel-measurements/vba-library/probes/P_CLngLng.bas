Attribute VB_Name = "P_CLngLng"
Option Explicit
' CLngLng (64-bit VBA only) (VBA.Conversion; MS-VBAL 6.1.2.x).
Private Const N As Long = 29

Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, c As Currency, dt As Date, ds As Double, sg As Single
    Dim l As Long, i As Integer, by As Byte, b As Boolean, s As String, o As Object
    On Error GoTo fail
#If Win64 Then
    Select Case k
        Case 1: v = CLngLng(0)
        Case 2: v = CLngLng(2 ^ 63)   ' Double 2^63 overflow
        Case 3: v = CLngLng(2147483648#)
        Case 4: v = CLngLng(-2147483649#)
        Case 5: v = CLngLng(2.5)
        Case 6: v = CLngLng(3.5)
        Case 7: v = CLngLng(-2.5)
        Case 8: v = CLngLng(4503599627370497#)   ' 2^52+1
        Case 9: v = CLngLng(-(2 ^ 63))   ' -2^63 exact
        Case 10: v = CLngLng(1E19)   ' overflow
        Case 11: v = CLngLng("9223372036854775807")   ' max string
        Case 12: v = CLngLng("9223372036854775808")   ' overflow
        Case 13: v = CLngLng("-9223372036854775808")   ' min string
        Case 14: v = CLngLng("12.5")
        Case 15: v = CLngLng("&HFFFFFFFFFFFFFFFF")   ' -1
        Case 16: v = CLngLng("&H7FFFFFFFFFFFFFFF")
        Case 17: v = CLngLng("&HFFFFFFFF")   ' Long -1 or 4294967295
        Case 18: v = CLngLng("&O17")
        Case 19: v = CLngLng("abc")   ' 13
        Case 20: v = CLngLng("")   ' 13
        Case 21: v = CLngLng(True)
        Case 22: v = CLngLng(Null)   ' 94
        Case 23: v = CLngLng(Empty)
        Case 24: v = CLngLng(CVErr(5))   ' 13
        Case 25: v = CLngLng(CCur(922337203685477.5807))
        Case 26: v = CLngLng(CDec("9223372036854775807"))
        Case 27: v = CLngLng(CDec("9223372036854775808"))   ' overflow
        Case 28: v = CLngLng(#1/2/2020#)
        Case 29: v = CLngLng(9223372036854775806 + 1)   ' LongLong arithmetic
    End Select
#Else
    v = "n/a-32bit"
#End If
    T = "CLngLng." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "CLngLng." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
