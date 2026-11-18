Attribute VB_Name = "P_FileCopy"
Option Explicit
' FileCopy (FileSystem; MS-VBAL 6.1.2.8.x). Files only under %TEMP%\cvp-lib\FileCopy.
Private Const N As Long = 20

Private Function Dr() As String
    Dr = Environ("TEMP") & "\cvp-lib\FileCopy"
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

Private Function Rd(ByVal p As String) As String
    Dim n As Integer, s As String
    n = FreeFile
    Open p For Binary Access Read As #n
    s = Space$(LOF(n))
    If LOF(n) > 0 Then Get #n, , s
    Close #n
    Rd = s
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String, e As Long, n As Integer, a As String, b As String
    On Error GoTo fail
    Wipe
    Mk
    a = Dr & "\a.txt"
    b = Dr & "\b.txt"
    Select Case k
        Case 1   ' basic copy: content equal, source kept
            Wr a, "hello"
            FileCopy a, b
            v = Rd(b) & "," & (Dir(a) = "a.txt")
        Case 2   ' same length
            Wr a, "hello world"
            FileCopy a, b
            v = FileLen(b)
        Case 3   ' existing destination is overwritten
            Wr a, "hello"
            Wr b, "old and much longer text"
            FileCopy a, b
            v = Rd(b)
        Case 4   ' read-only destination: error 70
            Wr a, "hello"
            Wr b, "old"
            SetAttr b, vbReadOnly
            FileCopy a, b
        Case 5   ' source missing: error 53
            FileCopy Dr & "\nope.txt", b
        Case 6   ' destination folder missing: error 76
            Wr a, "x"
            FileCopy a, Dr & "\nodir\b.txt"
        Case 7   ' source open for Output
            Wr a, "x"
            n = FreeFile
            Open a For Output As #n
            FileCopy a, b
        Case 8   ' source open for Input
            Wr a, "xyz"
            n = FreeFile
            Open a For Input As #n
            FileCopy a, b
            v = Rd(b)
        Case 9   ' attributes are copied
            Wr a, "x"
            SetAttr a, vbReadOnly + vbHidden
            FileCopy a, b
            v = GetAttr(b) And 7
            SetAttr a, vbNormal
            SetAttr b, vbNormal
        Case 10  ' wildcard in the source
            Wr a, "x"
            FileCopy Dr & "\*.txt", b
        Case 11  ' copy onto itself
            Wr a, "x"
            FileCopy a, a
        Case 12  ' empty file
            Wr a, ""
            FileCopy a, b
            v = FileLen(b)
        Case 13  ' destination is a folder
            Wr a, "x"
            MkDir Dr & "\sub"
            FileCopy a, Dr & "\sub"
        Case 14  ' source is a folder
            MkDir Dr & "\sub"
            FileCopy Dr & "\sub", b
        Case 15  ' binary content survives (NUL, Ctrl-Z, CRLF)
            Wr a, "a" & Chr$(0) & Chr$(26) & vbCrLf & "z"
            FileCopy a, b
            v = Sh(Rd(b))
        Case 16  ' empty source name
            FileCopy "", b
        Case 17  ' Null source: error 94
            FileCopy Null, b
        Case 18  ' named arguments
            Wr a, "named"
            FileCopy Source:=a, Destination:=b
            v = Rd(b)
        Case 19  ' modification time is kept (within two seconds)
            Wr a, "x"
            FileCopy a, b
            v = (Abs(FileDateTime(a) - FileDateTime(b)) < 2 / 86400)
        Case 20  ' empty destination name
            Wr a, "x"
            FileCopy a, ""
    End Select
    s = "FileCopy." & k & "=" & Sh(v)
    Wipe
    T = s
    Exit Function
fail:
    e = Err.Number
    Wipe
    T = "FileCopy." & k & "=ERR" & e
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
