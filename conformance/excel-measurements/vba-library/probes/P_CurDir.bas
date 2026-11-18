Attribute VB_Name = "P_CurDir"
Option Explicit
' CurDir / CurDir$ (FileSystem; MS-VBAL 6.1.2.8.x). Files only under %TEMP%\cvp-lib\CurDir.
Private Const N As Long = 15
Private sv As String

Private Function Dr() As String
    Dr = Environ("TEMP") & "\cvp-lib\CurDir"
End Function

Private Function Dv() As String
    Dv = Left$(Environ("TEMP"), 1)
End Function

Private Function EndsW(ByVal s As String, ByVal t As String) As Boolean
    EndsW = (LCase$(Right$(s, Len(t))) = LCase$(t))
End Function

Private Sub Mk()
    On Error Resume Next
    MkDir Environ("TEMP") & "\cvp-lib"
    MkDir Dr
End Sub

Private Sub WipeTree(ByVal d As String)
    Dim f As String, s As String, a() As String, i As Long, p As String
    On Error Resume Next
    If InStr(1, d, "\cvp-lib\", vbTextCompare) = 0 Then Exit Sub
    f = Dir(d & "\*", vbNormal + vbHidden + vbReadOnly + vbSystem + vbDirectory)
    Do While Len(f) > 0
        If f <> "." And f <> ".." Then s = s & f & "|"
        f = Dir()
    Loop
    If Len(s) > 0 Then
        a = Split(s, "|")
        For i = 0 To UBound(a)
            If Len(a(i)) > 0 Then
                p = d & "\" & a(i)
                SetAttr p, vbNormal
                If (GetAttr(p) And vbDirectory) <> 0 Then
                    WipeTree p
                Else
                    Kill p
                End If
            End If
        Next
    End If
    SetAttr d, vbNormal
    RmDir d
End Sub

Private Sub Wipe()
    On Error Resume Next
    Close
    If Len(sv) > 0 Then
        ChDrive Left$(sv, 1)
        ChDir sv
    End If
    WipeTree Dr
    RmDir Environ("TEMP") & "\cvp-lib"
End Sub

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String, e As Long
    On Error GoTo fail
    sv = CurDir
    Wipe
    Mk
    Select Case k
        Case 1: v = TypeName(CurDir)                                  ' Variant holding a String
        Case 2: v = TypeName(CurDir$)                                 ' $ form: String
        Case 3: v = (CurDir = CurDir$)                                ' both forms agree
        Case 4: ChDir Dr: v = EndsW(CurDir(Dv()), "\CurDir")          ' after ChDir
        Case 5: v = (CurDir(LCase$(Dv())) = CurDir(UCase$(Dv())))     ' letter case does not matter
        Case 6: v = (CurDir("") = CurDir)                             ' empty drive: current drive
        Case 7: v = (CurDir(Dv() & ":") = CurDir(Dv()))               ' multi-character: first letter only
        Case 8: v = CurDir("1")                                       ' not a drive: error
        Case 9: v = CurDir(Null)                                      ' Null drive
        Case 10: v = (CurDir(Empty) = CurDir)                         ' Empty drive
        Case 11: ChDir Left$(Environ("TEMP"), 3): v = Len(CurDir(Dv())) & ":" & Right$(CurDir(Dv()), 1)   ' root keeps its backslash
        Case 12: ChDir Dr: v = (Right$(CurDir(Dv()), 1) = "\")        ' non-root: no trailing backslash
        Case 13: v = TypeName(CurDir$(Dv()))                          ' $ form with drive
        Case 14: v = (CurDir(Drive:=Dv()) = CurDir(Dv()))             ' named argument
        Case 15: v = (CurDir$(Drive:=Dv()) = CurDir$(Dv()))           ' named argument, $ form
    End Select
    s = "CurDir." & k & "=" & Sh(v)
    Wipe
    T = s
    Exit Function
fail:
    e = Err.Number
    Wipe
    T = "CurDir." & k & "=ERR" & e
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
