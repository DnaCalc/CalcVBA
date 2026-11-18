Attribute VB_Name = "P_SetAttr"
Option Explicit
' SetAttr (FileSystem; MS-VBAL 6.1.2.8.x). Files only under %TEMP%\cvp-lib\SetAttr.
Private Const N As Long = 20
Private sv As String

Private Function Dr() As String
    Dr = Environ("TEMP") & "\cvp-lib\SetAttr"
End Function

Private Function Cd() As String
    Cd = CurDir(Left$(Environ("TEMP"), 1))
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
    Dim v As Variant, s As String, e As Long, n As Integer, f As String
    On Error GoTo fail
    sv = Cd()
    Wipe
    Mk
    f = Dr & "\f.txt"
    Wr f, "x"
    Select Case k
        Case 1: SetAttr f, vbReadOnly: v = GetAttr(f)                       ' read-only
        Case 2: SetAttr f, vbHidden: v = GetAttr(f)                         ' hidden
        Case 3: SetAttr f, vbSystem: v = GetAttr(f)                         ' system
        Case 4: SetAttr f, vbReadOnly: SetAttr f, vbNormal: v = GetAttr(f)  ' vbNormal clears
        Case 5: SetAttr f, vbReadOnly + vbHidden: v = GetAttr(f)            ' combination
        Case 6: MkDir Dr & "\sub": SetAttr Dr & "\sub", vbHidden: v = GetAttr(Dr & "\sub")   ' folder: directory bit stays
        Case 7: SetAttr f, vbDirectory: v = GetAttr(f)                      ' directory bit on a file
        Case 8: SetAttr f, vbArchive: v = GetAttr(f)                        ' archive
        Case 9: SetAttr Dr & "\nope.txt", vbReadOnly                        ' missing file: error 53
        Case 10: SetAttr Dr & "\nodir\x.txt", vbReadOnly                    ' missing folder
        Case 11   ' open file
            n = FreeFile
            Open f For Input As #n
            SetAttr f, vbReadOnly
            v = GetAttr(f)
        Case 12: SetAttr Null, vbNormal                                     ' Null name: error 94
        Case 13: SetAttr f, 1000: v = GetAttr(f)                            ' invalid attribute value
        Case 14: SetAttr f, vbVolume: v = GetAttr(f)                        ' volume bit
        Case 15: SetAttr PathName:=f, Attributes:=vbReadOnly: v = GetAttr(f)   ' named arguments
        Case 16: SetAttr f, -1: v = GetAttr(f)                              ' negative value
        Case 17: SetAttr Dr & "\*.txt", vbReadOnly                          ' wildcard
        Case 18   ' read-only file cannot be killed until cleared
            SetAttr f, vbReadOnly
            On Error Resume Next
            Kill f
            e = Err.Number
            On Error GoTo fail
            SetAttr f, vbNormal
            Kill f
            v = e & "," & (Dir(f) = "")
        Case 19: SetAttr f, "1": v = GetAttr(f)                             ' numeric string attribute
        Case 20: SetAttr Dr, vbReadOnly: v = GetAttr(Dr)                    ' folder read-only
    End Select
    s = "SetAttr." & k & "=" & Sh(v)
    Wipe
    T = s
    Exit Function
fail:
    e = Err.Number
    Wipe
    T = "SetAttr." & k & "=ERR" & e
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
