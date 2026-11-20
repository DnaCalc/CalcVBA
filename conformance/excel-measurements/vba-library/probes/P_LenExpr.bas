Attribute VB_Name = "P_LenExpr"
Option Explicit
' Len / LenB of typed expressions other than plain variables (Strings; MS-VBAL 6.1.2.11.1.x).
' P_Len measures variables; these settle whether the storage-size rule also covers expressions
' and whether a call inside the argument still runs.
Private Const N As Long = 14
Private calls As Long

Private Function F() As Long
    calls = calls + 1
    F = 12345
End Function

Private Function G() As Long
    calls = calls + 1
    G = 1
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant, lg As Long, ig As Integer, db As Double, arr(1 To 2) As Long
    On Error GoTo fail
    lg = 7: ig = 3: db = 1.5: arr(1) = 99
    calls = 0
    Select Case k
        Case 1: v = "compile-refused"
        Case 2: v = "compile-refused"
        Case 3: v = "compile-refused"
        Case 4: v = "compile-refused"
        Case 5: v = "compile-refused"
        Case 6: v = "compile-refused"
        Case 7: v = "compile-refused"
        Case 8: v = "compile-refused"
        Case 9: v = Len(arr(G())) & "|" & calls ' array element, index has a call
        Case 10: v = Len(arr(1))                ' Long array element
        Case 11: v = "compile-refused"
        Case 12: v = "compile-refused"
        Case 13: v = "compile-refused"
        Case 14: v = "compile-refused"
    End Select
    T = "LenExpr." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "LenExpr." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
