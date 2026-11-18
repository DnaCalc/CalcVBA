Attribute VB_Name = "P_ChDir"
Option Explicit
' ChDir (FileSystem; MS-VBAL 6.1.2.8.x). Files only under %TEMP%\cvp-lib\ChDir.
Private Const N As Long = 15
Private sv As String

Private Function Dr() As String
    Dr = Environ("TEMP") & "\cvp-lib\ChDir"
End Function

Private Function Dv() As String
    Dv = Left$(Environ("TEMP"), 1)
End Function

Private Function Cd() As String
    Cd = CurDir(Dv())
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
    If Len(sv) > 0 Then ChDir sv
    WipeTree Dr
    RmDir Environ("TEMP") & "\cvp-lib"
End Sub

Private Sub Wr(ByVal p As String, ByVal s As String)
    Dim n As Integer
    n = FreeFile
    Open p For Binary Access Write As #n
    Put #n, , s
    Close #n
End Sub

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String, e As Long, a As Boolean, b As Boolean
    On Error GoTo fail
    sv = Cd()
    Wipe
    Mk
    Select Case k
        Case 1: ChDir Dr: v = EndsW(Cd(), "\chdir")                ' absolute path
        Case 2: ChDir Dr & "\": v = EndsW(Cd(), "\chdir")          ' trailing backslash
        Case 3: MkDir Dr & "\sub": ChDir Dr: ChDir "sub": v = EndsW(Cd(), "\chdir\sub")   ' relative
        Case 4: MkDir Dr & "\sub": ChDir Dr & "\sub": ChDir "..": v = EndsW(Cd(), "\chdir")   ' parent
        Case 5: ChDir Dr & "\nope"                                 ' missing folder: error 76
        Case 6: ChDir ""                                           ' empty path
        Case 7: Wr Dr & "\f.txt", "x": ChDir Dr & "\f.txt"         ' a file, not a folder
        Case 8: ChDir Left$(Environ("TEMP"), 3): v = Len(Cd()) & ":" & (Right$(Cd(), 1) = "\")   ' drive root
        Case 9: ChDir Dr: ChDir Left$(Environ("TEMP"), 2): v = EndsW(Cd(), "\chdir")   ' drive only: no change
        Case 10: ChDir Replace(Dr, "\", "/"): v = EndsW(Cd(), "\chdir")                 ' forward slashes
        Case 11: ChDir Null                                        ' Null: error 94
        Case 12: ChDir Dr & "\*"                                   ' wildcard
        Case 13: ChDir LCase$(Dr): a = (Cd() = Dr): b = (Cd() = LCase$(Dr)): v = a & "," & b   ' case as given or as on disk
        Case 14: MkDir Dr & "\sub": ChDir Dr & "\sub ": v = EndsW(Cd(), "\sub")         ' trailing space
        Case 15: MkDir Dr & "\sub": ChDir Dr & "\sub": ChDir "..\..": v = EndsW(Cd(), "\cvp-lib")   ' two levels up
    End Select
    s = "ChDir." & k & "=" & Sh(v)
    Wipe
    T = s
    Exit Function
fail:
    e = Err.Number
    Wipe
    T = "ChDir." & k & "=ERR" & e
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
