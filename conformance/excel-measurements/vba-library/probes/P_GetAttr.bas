Attribute VB_Name = "P_GetAttr"
Option Explicit
' GetAttr (FileSystem; MS-VBAL 6.1.2.8.x). Files only under %TEMP%\cvp-lib\GetAttr.
Private Const N As Long = 20
Private sv As String

Private Function Dr() As String
    Dr = Environ("TEMP") & "\cvp-lib\GetAttr"
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
    Select Case k
        Case 1: Wr f, "x": v = GetAttr(f)                                     ' fresh file: archive bit
        Case 2: Wr f, "x": SetAttr f, vbNormal: v = GetAttr(f)                ' vbNormal = 0
        Case 3: Wr f, "x": SetAttr f, vbReadOnly: v = GetAttr(f)              ' read-only
        Case 4: Wr f, "x": SetAttr f, vbHidden: v = GetAttr(f)                ' hidden
        Case 5: Wr f, "x": SetAttr f, vbSystem: v = GetAttr(f)                ' system
        Case 6: v = GetAttr(Dr)                                               ' a folder
        Case 7: MkDir Dr & "\sub": SetAttr Dr & "\sub", vbHidden: v = GetAttr(Dr & "\sub")   ' hidden folder
        Case 8: Wr f, "x": SetAttr f, vbReadOnly + vbHidden + vbSystem: v = GetAttr(f)       ' combination
        Case 9: v = GetAttr(Left$(Environ("TEMP"), 3))                        ' drive root (machine specific)
        Case 10: v = GetAttr(Dr & "\nope.txt")                                ' missing file: error 53
        Case 11: v = GetAttr(Dr & "\nodir\x.txt")                             ' missing folder
        Case 12: v = GetAttr("")                                              ' empty name
        Case 13: v = GetAttr(Null)                                            ' Null: error 94
        Case 14: Wr f, "x": v = GetAttr(Dr & "\*.txt")                        ' wildcard
        Case 15: Wr f, "x": v = TypeName(GetAttr(f))                          ' result type
        Case 16: v = GetAttr(Dr & "\")                                        ' folder with trailing backslash
        Case 17: Wr f, "x": SetAttr f, vbReadOnly: v = ((GetAttr(f) And vbReadOnly) = vbReadOnly) & "," & ((GetAttr(f) And vbHidden) = vbHidden)   ' bit tests
        Case 18: Wr f, "x": v = GetAttr(PathName:=f)                          ' named argument
        Case 19   ' archive bit set again by a write
            Wr f, "x"
            SetAttr f, vbNormal
            n = FreeFile
            Open f For Append As #n
            Print #n, "y"
            Close #n
            v = GetAttr(f)
        Case 20   ' attributes of a file that is open
            Wr f, "x"
            SetAttr f, vbReadOnly
            n = FreeFile
            Open f For Input As #n
            v = GetAttr(f)
    End Select
    s = "GetAttr." & k & "=" & Sh(v)
    Wipe
    T = s
    Exit Function
fail:
    e = Err.Number
    Wipe
    T = "GetAttr." & k & "=ERR" & e
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
