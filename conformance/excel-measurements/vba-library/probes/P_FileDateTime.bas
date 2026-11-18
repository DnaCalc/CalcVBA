Attribute VB_Name = "P_FileDateTime"
Option Explicit
' FileDateTime (FileSystem; MS-VBAL 6.1.2.8.x). Files only under %TEMP%\cvp-lib\FileDateTime.
' File times are never printed: only types and relations to Now.
Private Const N As Long = 17
Private sv As String

Private Function Dr() As String
    Dr = Environ("TEMP") & "\cvp-lib\FileDateTime"
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
    Dim v As Variant, s As String, e As Long, n As Integer, f As String, d1 As Date, d2 As Date
    On Error GoTo fail
    sv = Cd()
    Wipe
    Mk
    f = Dr & "\f.txt"
    Select Case k
        Case 1: Wr f, "x": v = TypeName(FileDateTime(f))                      ' Date inside a Variant
        Case 2: Wr f, "x": v = (Abs(FileDateTime(f) - Now) < 1)               ' created just now
        Case 3: MkDir Dr & "\sub": Wr Dr & "\sub\g.txt", "x": v = (Abs(FileDateTime(Dr & "\sub\g.txt") - Now) < 1)   ' file in a subfolder
        Case 4: v = TypeName(FileDateTime(Dr)) & "," & (Abs(FileDateTime(Dr) - Now) < 1)    ' a folder
        Case 5: v = FileDateTime(Dr & "\nope.txt")                            ' missing file: error 53
        Case 6: v = FileDateTime(Dr & "\nodir\x.txt")                         ' missing folder
        Case 7: v = FileDateTime("")                                          ' empty name
        Case 8: v = FileDateTime(Null)                                        ' Null: error 94
        Case 9: Wr f, "x": v = FileDateTime(Dr & "\*.txt")                    ' wildcard
        Case 10: Wr f, "x": d1 = FileDateTime(f): v = (d1 <= Now + 1 / 86400)  ' not in the future
        Case 11   ' second file written later is not older
            Wr f, "x"
            d1 = FileDateTime(f)
            Wr Dr & "\g.txt", "y"
            d2 = FileDateTime(Dr & "\g.txt")
            v = (d2 >= d1)
        Case 12: Wr f, "x": v = TypeName(FileDateTime(PathName:=f))           ' named argument
        Case 13: n = FreeFile: Open f For Output As #n: Print #n, "abc": v = TypeName(FileDateTime(f))   ' file open for writing
        Case 14: v = TypeName(FileDateTime(Dr & "\"))                         ' folder with trailing backslash
        Case 15: Wr f, "x": ChDir Dr: v = (Abs(FileDateTime("f.txt") - Now) < 1)   ' relative name
        Case 16: v = FileDateTime(vbNullString)                               ' vbNullString
        Case 17: v = FileDateTime(123)                                        ' numeric argument: file "123" missing
    End Select
    s = "FileDateTime." & k & "=" & Sh(v)
    Wipe
    T = s
    Exit Function
fail:
    e = Err.Number
    Wipe
    T = "FileDateTime." & k & "=ERR" & e
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
