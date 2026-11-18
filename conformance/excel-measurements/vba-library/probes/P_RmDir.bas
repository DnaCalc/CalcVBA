Attribute VB_Name = "P_RmDir"
Option Explicit
' RmDir (FileSystem; MS-VBAL 6.1.2.8.x). Files only under %TEMP%\cvp-lib\RmDir.
Private Const N As Long = 18
Private sv As String

Private Function Dr() As String
    Dr = Environ("TEMP") & "\cvp-lib\RmDir"
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

Private Function Has(ByVal p As String) As Boolean
    Has = (Len(Dir(p, vbDirectory + vbHidden + vbSystem)) > 0)
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String, e As Long, n As Integer
    On Error GoTo fail
    sv = Cd()
    Wipe
    Mk
    Select Case k
        Case 1: MkDir Dr & "\n1": RmDir Dr & "\n1": v = Has(Dr & "\n1")           ' empty folder removed
        Case 2: MkDir Dr & "\n1": Wr Dr & "\n1\f.txt", "x": RmDir Dr & "\n1"       ' not empty: error 75
        Case 3: RmDir Dr & "\nope"                                                 ' missing: error 76
        Case 4: RmDir ""                                                           ' empty name
        Case 5: RmDir Null                                                         ' Null: error 94
        Case 6: MkDir Dr & "\n1": ChDir Dr & "\n1": RmDir Dr & "\n1"               ' current folder
        Case 7: MkDir Dr & "\n1": RmDir Dr & "\n1\": v = Has(Dr & "\n1")           ' trailing backslash
        Case 8: Wr Dr & "\f.txt", "x": RmDir Dr & "\f.txt"                         ' a file, not a folder
        Case 9: MkDir Dr & "\n1": SetAttr Dr & "\n1", vbReadOnly: RmDir Dr & "\n1": v = Has(Dr & "\n1")   ' read-only folder
        Case 10: MkDir Dr & "\n1": RmDir Dr & "\n*"                                ' wildcard
        Case 11: MkDir Dr & "\n1": MkDir Dr & "\n1\n2": RmDir Dr & "\n1\n2": RmDir Dr & "\n1": v = Has(Dr & "\n1")   ' nested, inner first
        Case 12: MkDir Dr & "\n1": ChDir Dr: RmDir "n1": v = Has(Dr & "\n1")      ' relative name
        Case 13: MkDir Dr & "\n1": RmDir Path:=Dr & "\n1": v = Has(Dr & "\n1")    ' named argument
        Case 14: MkDir Dr & "\Abc": RmDir UCase$(Dr) & "\ABC": v = Has(Dr & "\Abc")   ' case does not matter
        Case 15: MkDir Dr & "\n1": Wr Dr & "\n1\h.txt", "x": SetAttr Dr & "\n1\h.txt", vbHidden: RmDir Dr & "\n1"   ' only a hidden file inside: error 75
        Case 16: MkDir Dr & "\n1": RmDir Dr & "\n1": RmDir Dr & "\n1"              ' second RmDir: error 76
        Case 17   ' a file is open inside
            MkDir Dr & "\n1"
            n = FreeFile
            Open Dr & "\n1\f.txt" For Output As #n
            RmDir Dr & "\n1"
        Case 18: MkDir Dr & "\n1": RmDir Dr & "\n1\..\n1": v = Has(Dr & "\n1")    ' path with ..
    End Select
    s = "RmDir." & k & "=" & Sh(v)
    Wipe
    T = s
    Exit Function
fail:
    e = Err.Number
    Wipe
    T = "RmDir." & k & "=ERR" & e
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
