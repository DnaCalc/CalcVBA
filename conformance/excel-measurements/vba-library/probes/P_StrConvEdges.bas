Attribute VB_Name = "P_StrConvEdges"
Option Explicit
' Combinations, Byte() source coercion, LCID validation and casing captured on ACP 1252.
Private Const N As Long = 18
Private Function T(ByVal k As Long) As String
    Dim v As Variant, b(0 To 3) As Byte
    On Error GoTo fail
    b(0) = 65: b(1) = 66: b(2) = 67: b(3) = 0
    Select Case k
        Case 1: v = StrConv("abc", 65)
        Case 2: v = StrConv("ABC", 66)
        Case 3: v = StrConv("ABC", 67)
        Case 4: v = StrConv("abc", 129)
        Case 5: v = StrConv("ABC", 130)
        Case 6: v = StrConv("ABC", 131)
        Case 7: v = StrConv("abc", 1073741824)
        Case 8: v = StrConv("abc", 1, 66567)
        Case 9: v = StrConv("abc", 1, -1)
        Case 10: v = StrConv(b, 0)
        Case 11: v = StrConv(b, 1)
        Case 12: v = StrConv(b, 3)
        Case 13: v = StrConv("a" & ChrW(&HA0) & "b", 3)
        Case 14: v = StrConv("a" & ChrW(&H3000) & "b", 3)
        Case 15: v = StrConv(ChrW(&H130), 2)
        Case 16: v = StrConv(ChrW(&HD800), 1)
        Case 17: v = StrConv(Null, Null)
        Case 18: v = StrConv("abc", 256)
    End Select
    T = "StrConvEdges." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "StrConvEdges." & k & "=ERR" & Err.Number
End Function
Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
