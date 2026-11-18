Attribute VB_Name = "P_FileLen"
Option Explicit
' FileLen (FileSystem; MS-VBAL 6.1.2.8.x). Files only under %TEMP%\cvp-lib\FileLen.
Private Const N As Long = 19
Private sv As String

Private Function Dr() As String
    Dr = Environ("TEMP") & "\cvp-lib\FileLen"
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
        Case 1: Wr f, "hello": v = FileLen(f)                                 ' five bytes
        Case 2: Wr f, "": v = FileLen(f)                                      ' empty file
        Case 3: Wr f, "x": v = TypeName(FileLen(f))                           ' Long
        Case 4   ' file open for Output with unflushed Print
            n = FreeFile
            Open f For Output As #n
            Print #n, "abc"
            v = FileLen(f)
        Case 5   ' after Print and Close: text plus CRLF
            n = FreeFile
            Open f For Output As #n
            Print #n, "abc"
            Close #n
            v = FileLen(f)
        Case 6: Wr f, String$(3000, "x"): v = FileLen(f)                      ' 3000 bytes
        Case 7: v = FileLen(Dr)                                               ' a folder
        Case 8: v = FileLen(Dr & "\nope.txt")                                 ' missing: error 53
        Case 9: v = FileLen("")                                               ' empty name
        Case 10: v = FileLen(Null)                                            ' Null: error 94
        Case 11: Wr f, "x": v = FileLen(Dr & "\*.txt")                        ' wildcard
        Case 12: Wr f, "abc": ChDir Dr: v = FileLen("f.txt")                  ' relative name
        Case 13: Wr f, "abc": v = FileLen(PathName:=f)                        ' named argument
        Case 14   ' Append adds to the length
            Wr f, "abcd"
            n = FreeFile
            Open f For Append As #n
            Print #n, "xy";
            Close #n
            v = FileLen(f)
        Case 15: Wr f, String$(40000, "x"): v = FileLen(f) & ":" & TypeName(FileLen(f))   ' beyond Integer
        Case 16   ' equals LOF while open in Binary mode
            Wr f, "abcdefg"
            n = FreeFile
            Open f For Binary As #n
            v = (FileLen(f) = LOF(n))
        Case 17: v = FileLen(vbNullString)                                    ' vbNullString
        Case 18: v = FileLen(Dr & "\nodir\x.txt")                             ' missing folder: error 76
        Case 19: Wr f, "abc": v = FileLen(UCase$(f))                          ' case does not matter
    End Select
    s = "FileLen." & k & "=" & Sh(v)
    Wipe
    T = s
    Exit Function
fail:
    e = Err.Number
    Wipe
    T = "FileLen." & k & "=ERR" & e
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
