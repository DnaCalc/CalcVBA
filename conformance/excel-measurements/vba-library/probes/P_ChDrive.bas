Attribute VB_Name = "P_ChDrive"
Option Explicit
' ChDrive (FileSystem; MS-VBAL 6.1.2.8.x). Only the TEMP folder's drive is used.
Private Const N As Long = 13
Private sv As String

Private Function Dr() As String
    Dr = Environ("TEMP") & "\cvp-lib\ChDrive"
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
    If Len(sv) > 0 Then
        ChDrive Left$(sv, 1)
        ChDir sv
    End If
    WipeTree Dr
    RmDir Environ("TEMP") & "\cvp-lib"
End Sub

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String, e As Long, before As String
    On Error GoTo fail
    sv = CurDir
    Wipe
    Mk
    before = CurDir
    Select Case k
        Case 1: ChDrive Dv(): v = (UCase$(Left$(CurDir, 2)) = UCase$(Dv()) & ":")       ' the TEMP drive letter
        Case 2: ChDrive LCase$(Dv()): v = (UCase$(Left$(CurDir, 2)) = UCase$(Dv()) & ":")   ' lower case letter
        Case 3: ChDrive Dv() & ":": v = (UCase$(Left$(CurDir, 2)) = UCase$(Dv()) & ":")     ' letter and colon
        Case 4: ChDrive Dv() & "ZZZ": v = (UCase$(Left$(CurDir, 2)) = UCase$(Dv()) & ":")   ' multi-character: first letter only
        Case 5: ChDrive "": v = (CurDir = before)                                        ' empty string: no change
        Case 6: ChDrive "1"                                                              ' not a drive letter
        Case 7: ChDrive Null                                                             ' Null: error 94
        Case 8: ChDrive 5                                                                ' numeric argument
        Case 9: ChDrive Dv(): v = (CurDir = before)                                      ' same drive: current dir unchanged
        Case 10: ChDir Dr: ChDrive Dv(): v = EndsW(Cd(), "\ChDrive")                      ' ChDrive keeps the drive's directory
        Case 11: ChDrive " " & Dv()                                                      ' leading space
        Case 12: ChDrive "\"                                                             ' backslash
        Case 13: ChDrive vbNullString: v = (CurDir = before)                              ' vbNullString
    End Select
    s = "ChDrive." & k & "=" & Sh(v)
    Wipe
    T = s
    Exit Function
fail:
    e = Err.Number
    Wipe
    T = "ChDrive." & k & "=ERR" & e
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
