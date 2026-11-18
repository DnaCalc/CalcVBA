Attribute VB_Name = "P_ShellEdges"
Option Explicit
' UNMEASURED style/coercion edges; all executable names remain the recorded missing path.
Private Const NOEXE As String = "C:\NoSuchDir_CalcVbaProbe\nosuch.exe"
Private Const N As Long = 15
Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Shell(NOEXE, 8)
        Case 2: v = Shell(NOEXE, 9)
        Case 3: v = Shell(NOEXE, 10)
        Case 4: v = Shell(NOEXE, 11)
        Case 5: v = Shell(NOEXE, 12)
        Case 6: v = Shell(NOEXE, True)
        Case 7: v = Shell(NOEXE, Empty)
        Case 8: v = Shell(NOEXE, Null)
        Case 9: v = Shell(NOEXE, "0")
        Case 10: v = Shell(NOEXE, "abc")
        Case 11: v = Shell(NOEXE, 1.5)
        Case 12: v = Shell(NOEXE, 2.5)
        Case 13: v = Shell(NOEXE, "")
        Case 14: v = Shell("   ")
        Case 15: v = Shell(vbTab & NOEXE, 0)
    End Select
    T = "ShellEdges." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "ShellEdges." & k & "=ERR" & Err.Number
End Function
Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
