Attribute VB_Name = "P_Kill"
Option Explicit
' Kill (FileSystem; MS-VBAL 6.1.2.8.x). Files only under %TEMP%\cvp-lib\Kill.
Private Const N As Long = 18

Private Function Dr() As String
    Dr = Environ("TEMP") & "\cvp-lib\Kill"
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

' Names left in the folder (files and folders, any attribute), sorted by Dir's order.
Private Function Left1() As String
    Dim f As String, s As String
    f = Dir(Dr & "\*", vbNormal + vbHidden + vbReadOnly + vbSystem + vbDirectory)
    Do While Len(f) > 0
        If f <> "." And f <> ".." Then s = s & f & "|"
        f = Dir()
    Loop
    Left1 = s
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String, e As Long, n As Integer, f As String
    On Error GoTo fail
    Wipe
    Mk
    f = Dr & "\f.txt"
    Select Case k
        Case 1: Wr f, "x": Kill f: v = "[" & Left1() & "]"                    ' single file removed
        Case 2: Kill Dr & "\nope.txt"                                         ' missing: error 53
        Case 3   ' wildcard removes all matches
            Wr Dr & "\a.txt", "a": Wr Dr & "\b.txt", "b": Wr Dr & "\c.dat", "c"
            Kill Dr & "\*.txt"
            v = Left1()
        Case 4   ' ? wildcard
            Wr Dr & "\a1.txt", "a": Wr Dr & "\a22.txt", "b": Wr Dr & "\b1.txt", "c"
            Kill Dr & "\a?.txt"
            v = Left1()
        Case 5   ' read-only file: error 70
            Wr f, "x": SetAttr f, vbReadOnly
            Kill f
        Case 6   ' hidden file
            Wr f, "x": SetAttr f, vbHidden
            Kill f
            v = "[" & Left1() & "]"
        Case 7   ' system file
            Wr f, "x": SetAttr f, vbSystem
            Kill f
            v = "[" & Left1() & "]"
        Case 8   ' file open for Output
            n = FreeFile
            Open f For Output As #n
            Kill f
        Case 9: MkDir Dr & "\sub": Kill Dr & "\sub"                           ' a folder
        Case 10: Kill Dr & "\nodir\x.txt"                                     ' missing folder: error 76
        Case 11: Wr f, "x": Kill Dr & "\*.zzz"                                ' wildcard without match
        Case 12: Kill ""                                                      ' empty name
        Case 13: Kill Null                                                    ' Null: error 94
        Case 14   ' * leaves folders alone
            Wr Dr & "\a.txt", "a"
            MkDir Dr & "\sub"
            Kill Dr & "\*"
            v = Left1()
        Case 15: Wr f, "x": Kill PathName:=f: v = "[" & Left1() & "]"         ' named argument
        Case 16: Wr f, "x": Kill f: Kill f                                    ' second Kill: error 53
        Case 17: Kill Dr & "\" & 123                                          ' numeric name part: file missing
        Case 18   ' *.* with a read-only file among others
            Wr Dr & "\a.txt", "a": Wr Dr & "\b.txt", "b": Wr Dr & "\c.txt", "c"
            SetAttr Dr & "\b.txt", vbReadOnly
            On Error Resume Next
            Kill Dr & "\*.*"
            e = Err.Number
            On Error GoTo fail
            v = e & ":" & Left1()
    End Select
    s = "Kill." & k & "=" & Sh(v)
    Wipe
    T = s
    Exit Function
fail:
    e = Err.Number
    Wipe
    T = "Kill." & k & "=ERR" & e
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
