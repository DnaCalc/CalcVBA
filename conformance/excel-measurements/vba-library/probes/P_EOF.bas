Attribute VB_Name = "P_EOF"
Option Explicit
' EOF (FileSystem; MS-VBAL 6.1.2.8.x). Files only under %TEMP%\cvp-lib\EOF.
Private Const N As Long = 18

Private Function Dr() As String
    Dr = Environ("TEMP") & "\cvp-lib\EOF"
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
    Dim v As Variant, s As String, e As Long, n As Integer, f As String
    Dim l As String, x As Long, y As Long, b As Byte
    On Error GoTo fail
    Wipe
    Mk
    f = Dr & "\f.txt"
    n = FreeFile
    Select Case k
        Case 1   ' empty file, Input mode: True at once
            Wr f, ""
            Open f For Input As #n
            v = EOF(n)
        Case 2   ' Line Input through two lines
            Wr f, "line1" & vbCrLf & "line2"
            Open f For Input As #n
            s = EOF(n) & ","
            Line Input #n, l: s = s & EOF(n) & ","
            Line Input #n, l: s = s & EOF(n)
            v = s
        Case 3   ' trailing CRLF: EOF true after the only line
            Wr f, "a" & vbCrLf
            Open f For Input As #n
            Line Input #n, l
            v = EOF(n)
        Case 4   ' Ctrl-Z in the middle of sequential input
            Wr f, "ab" & Chr$(26) & "cd"
            Open f For Input As #n
            Line Input #n, l
            v = Sh(l) & "," & EOF(n)
        Case 5   ' Binary: false until the last byte is read
            Wr f, "abcd"
            Open f For Binary As #n
            s = EOF(n) & ","
            Get #n, , b: s = s & EOF(n) & ","
            Get #n, , b: Get #n, , b: Get #n, , b: s = s & EOF(n)
            v = s
        Case 6   ' Output mode: error 54
            Open f For Output As #n
            v = EOF(n)
        Case 7   ' Append mode: error 54
            Open f For Append As #n
            v = EOF(n)
        Case 8   ' Random, two 4-byte records
            Open f For Random As #n Len = 4
            x = 1: Put #n, 1, x
            x = 2: Put #n, 2, x
            Close #n
            n = FreeFile
            Open f For Random As #n Len = 4
            s = EOF(n) & ","
            Get #n, 1, y: s = s & EOF(n) & ","
            Get #n, 2, y: s = s & EOF(n)
            v = s
        Case 9   ' bad file numbers: 5 never opened, 0
            On Error Resume Next
            v = EOF(5)
            s = "5:" & Err.Number
            Err.Clear
            v = EOF(0)
            s = s & ",0:" & Err.Number
            On Error GoTo fail
            v = s
        Case 10  ' result type
            Wr f, "x"
            Open f For Input As #n
            v = TypeName(EOF(n))
        Case 11  ' after Close: error 52
            Wr f, "x"
            Open f For Input As #n
            Close #n
            v = EOF(n)
        Case 12  ' empty file in Binary mode
            Wr f, ""
            Open f For Binary As #n
            v = EOF(n)
        Case 13  ' Input # of numbers
            Wr f, "1,2"
            Open f For Input As #n
            Input #n, x: s = x & ":" & EOF(n) & ","
            Input #n, x: s = s & x & ":" & EOF(n)
            v = s
        Case 14  ' Binary Seek to last byte and past it
            Wr f, "abcd"
            Open f For Binary As #n
            Seek #n, 4: s = EOF(n) & ","
            Seek #n, 5: s = s & EOF(n)
            v = s
        Case 15  ' named argument
            Wr f, "x"
            Open f For Input As #n
            v = EOF(FileNumber:=n)
        Case 16  ' Null: error 94
            v = EOF(Null)
        Case 17  ' Long file number
            Wr f, "x"
            Open f For Input As #n
            x = n
            v = EOF(x)
        Case 18  ' Binary file ending in Ctrl-Z: not EOF until the real end
            Wr f, "ab" & Chr$(26)
            Open f For Binary As #n
            Get #n, , b: Get #n, , b: s = EOF(n) & ","
            Get #n, , b: s = s & EOF(n)
            v = s
    End Select
    s = "EOF." & k & "=" & Sh(v)
    Wipe
    T = s
    Exit Function
fail:
    e = Err.Number
    Wipe
    T = "EOF." & k & "=ERR" & e
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
