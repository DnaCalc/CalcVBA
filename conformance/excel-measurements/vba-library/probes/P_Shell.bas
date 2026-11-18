Attribute VB_Name = "P_Shell"
Option Explicit
' Shell (Interaction; only a missing executable or an invalid window style, never launches anything).
Private Const NOEXE As String = "C:\NoSuchDir_CalcVbaProbe\nosuch.exe"
Private Const N As Long = 19

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Shell(NOEXE)                                                          ' missing file: error 53
        Case 2: v = Shell(NOEXE, vbHide)                                                  ' hidden
        Case 3: v = Shell(NOEXE, vbNormalFocus)
        Case 4: v = Shell(NOEXE, vbMinimizedFocus)
        Case 5: v = Shell(NOEXE, vbMaximizedFocus)
        Case 6: v = Shell(NOEXE, vbNormalNoFocus)
        Case 7: v = Shell(NOEXE, vbMinimizedNoFocus)
        Case 8: v = Shell(NOEXE, 99)                                                      ' bad window style: error 5
        Case 9: v = Shell(NOEXE, -1)                                                      ' negative window style
        Case 10: v = Shell(NOEXE, 5)                                                       ' style 5 is not a VbAppWinStyle
        Case 11: v = Shell(NOEXE, 7)                                                       ' style 7
        Case 12: v = Shell("")                                                             ' empty command
        Case 13: v = Shell(Null)                                                           ' Null command
        Case 14: v = Shell(NOEXE & " arg1 arg2")                                           ' with arguments
        Case 15: v = Shell("""" & NOEXE & """ arg1")                                       ' quoted path
        Case 16: v = Shell("nosuchprogram_calcvbaprobe.exe")                               ' bare missing name
        Case 17: v = Shell(PathName:=NOEXE, WindowStyle:=vbHide)                           ' named arguments
        Case 18: v = vbHide & "," & vbNormalFocus & "," & vbMinimizedFocus & "," & vbMaximizedFocus & "," & vbNormalNoFocus & "," & vbMinimizedNoFocus   ' constants
        Case 19: v = Shell(NOEXE, 32768)                                                   ' style beyond Integer
    End Select
    T = "Shell." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Shell." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
