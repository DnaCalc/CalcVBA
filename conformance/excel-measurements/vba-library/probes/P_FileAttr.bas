Attribute VB_Name = "P_FileAttr"
Option Explicit
' FileAttr (FileSystem; MS-VBAL 6.1.2.8.x). Files only under %TEMP%\cvp-lib\FileAttr.
Private Const N As Long = 18

Private Function Dr() As String
    Dr = Environ("TEMP") & "\cvp-lib\FileAttr"
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

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String, e As Long, n As Integer, m As Integer, f As String
    On Error GoTo fail
    Wipe
    Mk
    f = Dr & "\f.txt"
    Wr f, "abc"
    n = FreeFile
    Select Case k
        Case 1: Open f For Input As #n: v = FileAttr(n, 1)                ' Input = 1
        Case 2: Open f For Output As #n: v = FileAttr(n, 1)               ' Output = 2
        Case 3: Open f For Append As #n: v = FileAttr(n, 1)               ' Append = 8
        Case 4: Open f For Binary As #n: v = FileAttr(n, 1)               ' Binary = 32
        Case 5: Open f For Random As #n Len = 4: v = FileAttr(n, 1)       ' Random = 4
        Case 6: Open f For Binary As #n: v = FileAttr(n)                  ' ReturnType omitted: the mode
        Case 7: Open f For Binary As #n: v = TypeName(FileAttr(n, 1)) & "," & (FileAttr(n, 2) <> 0) & "," & TypeName(FileAttr(n, 2))   ' type; OS handle is non-zero
        Case 8: Open f For Binary As #n: v = FileAttr(n, 3)               ' bad ReturnType: error 5
        Case 9: Open f For Binary As #n: v = FileAttr(n, 0)               ' ReturnType 0: error 5
        Case 10: v = FileAttr(7, 1)                                       ' never opened: error 52
        Case 11: Open f For Binary As #n: Close #n: v = FileAttr(n, 1)    ' after Close: error 52
        Case 12: Open f For Binary Access Read As #n: v = FileAttr(n, 1)  ' Access does not change the mode
        Case 13: Open f For Random Access Read Shared As #n Len = 2: v = FileAttr(n, 1)   ' Random with Access and Shared
        Case 14: Open f For Input As #n: v = FileAttr(FileNumber:=n, ReturnType:=1)   ' named arguments
        Case 15: Open f For Input As #n: v = FileAttr(n, "1")             ' numeric string ReturnType
        Case 16: Open f For Input As #n: v = FileAttr(n, Null)            ' Null ReturnType: error 94
        Case 17: v = FileAttr(0, 1)                                       ' file number 0: error 52
        Case 18   ' two files in two modes at once
            Open f For Input As #n
            m = FreeFile
            Open Dr & "\g.txt" For Output As #m
            v = FileAttr(n, 1) & "," & FileAttr(m, 1)
    End Select
    s = "FileAttr." & k & "=" & Sh(v)
    Wipe
    T = s
    Exit Function
fail:
    e = Err.Number
    Wipe
    T = "FileAttr." & k & "=ERR" & e
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
