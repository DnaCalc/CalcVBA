Attribute VB_Name = "P_InStrLcid"
Option Explicit
Option Compare Binary
' Excel 16.0.20527.20196 explicit LCID controls for InStr and InStrRev.
Private Const N As Long = 18

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = InStr(1, "I", ChrW(&H131), 1033)
        Case 2: v = InStrRev("I", ChrW(&H131), -1, 1033)
        Case 3: v = InStr(1, ChrW(&H130), "i", 1033)
        Case 4: v = InStrRev(ChrW(&H130), "i", -1, 1033)
        Case 5: v = InStr(1, "ABC", "b", 1033)
        Case 6: v = InStrRev("ABC", "b", -1, 1033)
        Case 7: v = InStr(1, "I", ChrW(&H131), 1055)
        Case 8: v = InStrRev("I", ChrW(&H131), -1, 1055)
        Case 9: v = InStr(1, ChrW(&H130), "i", 1055)
        Case 10: v = InStrRev(ChrW(&H130), "i", -1, 1055)
        Case 11: v = InStr(1, "ABC", "b", 1055)
        Case 12: v = InStrRev("ABC", "b", -1, 1055)
        Case 13: v = InStr(1, "I", ChrW(&H131), 65535)
        Case 14: v = InStrRev("I", ChrW(&H131), -1, 65535)
        Case 15: v = InStr(1, ChrW(&H130), "i", 65535)
        Case 16: v = InStrRev(ChrW(&H130), "i", -1, 65535)
        Case 17: v = InStr(1, "ABC", "b", 65535)
        Case 18: v = InStrRev("ABC", "b", -1, 65535)
    End Select
    T = "InStrLcid." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "InStrLcid." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
