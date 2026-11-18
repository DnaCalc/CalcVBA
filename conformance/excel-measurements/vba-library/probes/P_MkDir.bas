Attribute VB_Name = "P_MkDir"
Option Explicit
' MkDir (FileSystem; MS-VBAL 6.1.2.8.x). Files only under %TEMP%\cvp-lib\MkDir.
Private Const N As Long = 18
Private sv As String

Private Function Dr() As String
    Dr = Environ("TEMP") & "\cvp-lib\MkDir"
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
    Dim v As Variant, s As String, e As Long
    On Error GoTo fail
    sv = Cd()
    Wipe
    Mk
    Select Case k
        Case 1: MkDir Dr & "\n1": v = Dir(Dr & "\n1", vbDirectory) & "," & GetAttr(Dr & "\n1")   ' created, attribute 16
        Case 2: MkDir Dr & "\n1": MkDir Dr & "\n1"                           ' already exists: error 75
        Case 3: MkDir Dr & "\p\q"                                            ' parent missing: error 76
        Case 4: MkDir Dr & "\n1\": v = Dir(Dr & "\n1", vbDirectory)          ' trailing backslash
        Case 5: ChDir Dr: MkDir "rel": v = Dir(Dr & "\rel", vbDirectory)     ' relative name
        Case 6: MkDir ""                                                     ' empty name
        Case 7: MkDir Null                                                   ' Null: error 94
        Case 8: MkDir Dr & "\a<b"                                            ' invalid character
        Case 9: Wr Dr & "\f", "x": MkDir Dr & "\f"                           ' name of an existing file
        Case 10: MkDir Replace(Dr, "\", "/") & "/fw": v = Dir(Dr & "\fw", vbDirectory)   ' forward slashes
        Case 11: MkDir Dr & "\dot.": v = Dir(Dr & "\dot", vbDirectory)       ' trailing dot is dropped
        Case 12: MkDir Dr & "\" & 123: v = Dir(Dr & "\123", vbDirectory)     ' numeric name part
        Case 13: MkDir Path:=Dr & "\named": v = Dir(Dr & "\named", vbDirectory)   ' named argument
        Case 14: MkDir Dr & "\" & String$(100, "z"): v = Len(Dir(Dr & "\" & String$(100, "z"), vbDirectory))   ' long name
        Case 15: MkDir Dr & "\two words": v = Dir(Dr & "\two words", vbDirectory)  ' space in name
        Case 16: MkDir Dr & "\Abc": MkDir UCase$(Dr) & "\ABC"                ' existing, different case: error 75
        Case 17   ' new folder is empty
            MkDir Dr & "\n1"
            v = Dir(Dr & "\n1\*")
        Case 18: MkDir Dr & "\a?b"                                           ' wildcard character in the name
    End Select
    s = "MkDir." & k & "=" & Sh(v)
    Wipe
    T = s
    Exit Function
fail:
    e = Err.Number
    Wipe
    T = "MkDir." & k & "=ERR" & e
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
