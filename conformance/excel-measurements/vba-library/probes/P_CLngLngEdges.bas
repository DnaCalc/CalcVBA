Attribute VB_Name = "P_CLngLngEdges"
Option Explicit
' UNMEASURED target-width radix, exact fractional strings and endpoint rounding.
Private Const N As Long = 25
Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = CLngLng("&HFFFF")
        Case 2: v = CLngLng("&H80000000")
        Case 3: v = CLngLng("&H100000000")
        Case 4: v = CLngLng("&H8000000000000000")
        Case 5: v = CLngLng("&H10000000000000000")
        Case 6: v = CLngLng("-&HFFFFFFFF")
        Case 7: v = CLngLng("-&H8000000000000000")
        Case 8: v = CLngLng("&HFFFF&")
        Case 9: v = CLngLng("&O37777777777")
        Case 10: v = CLngLng("&O1777777777777777777777")
        Case 11: v = CLngLng("9223372036854775806.5")
        Case 12: v = CLngLng("9223372036854775807.5")
        Case 13: v = CLngLng("-9223372036854775808.5")
        Case 14: v = CLngLng("-9223372036854775809")
        Case 15: v = CLngLng("9.223372036854775807E18")
        Case 16: v = CLngLng("9.223372036854775807D18")
        Case 17: v = CLngLng("(9223372036854775808)")
        Case 18: v = CLngLng("9223372036854775808-")
        Case 19: v = CLngLng(CCur("1.5000"))
        Case 20: v = CLngLng(CCur("2.5000"))
        Case 21: v = CLngLng(CCur("-2.5000"))
        Case 22: v = CLngPtr("&HFFFFFFFF")
        Case 23: v = CLngLng(CCur("922337203685477.4999"))
        Case 24: v = CLngLng(CCur("922337203685477.5000"))
        Case 25: v = CLngLng(CCur("-922337203685477.4999"))
    End Select
    T = "CLngLngEdges." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "CLngLngEdges." & k & "=ERR" & Err.Number
End Function
Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
