Attribute VB_Name = "P_Reset"
Option Explicit
' Reset (FileSystem; MS-VBAL 6.1.2.8.x). Files only under %TEMP%\cvp-lib\Reset.
Private Const N As Long = 14

Private Function Dr() As String
    Dr = Environ("TEMP") & "\cvp-lib\Reset"
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
    Reset
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

' How many of the file numbers 1..255 are open.
Private Function OpenCount() As Long
    Dim i As Long, c As Long, m As Long
    On Error Resume Next
    For i = 1 To 255
        Err.Clear
        m = FileAttr(i, 1)
        If Err.Number = 0 Then c = c + 1
    Next
    OpenCount = c
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String, e As Long, n As Integer, f As String, g As String
    On Error GoTo fail
    Wipe
    Mk
    f = Dr & "\f.txt"
    g = Dr & "\g.txt"
    Wr f, "abc"
    Wr g, "def"
    Select Case k
        Case 1   ' three files open, then none
            Open f For Input Shared As #1
            Open f For Input Shared As #2
            Open g For Input As #3
            s = OpenCount() & ","
            Reset
            v = s & OpenCount()
        Case 2   ' EOF on a reset file number: error 52
            Open f For Input As #1
            Reset
            v = EOF(1)
        Case 3: Reset: v = OpenCount()                                      ' nothing open: no error
        Case 4   ' buffered Print reaches the file
            Open Dr & "\o.txt" For Output As #1
            Print #1, "abc"
            Reset
            v = FileLen(Dr & "\o.txt")
        Case 5: Reset: Reset: v = "ok"                                      ' twice in a row
        Case 6   ' high numbers are closed too
            Open f For Input As #256
            s = FileAttr(256, 1) & ","
            Reset
            v = s & FreeFile(1)
        Case 7   ' the same number can be used again
            Open f For Input As #1
            Reset
            Open g For Input As #1
            v = Input(3, #1)
        Case 8   ' LOF and Loc after Reset: error 52
            Open f For Binary As #1
            Reset
            v = LOF(1)
        Case 9   ' Binary Put reaches the file
            Open Dr & "\b.dat" For Binary As #1
            s = "abcdef"
            Put #1, , s
            Reset
            v = FileLen(Dr & "\b.dat")
        Case 10  ' Reset releases an Output file so it can be deleted
            Open Dr & "\o.txt" For Output As #1
            Reset
            Kill Dr & "\o.txt"
            v = Dir(Dr & "\o.txt") = ""
        Case 11  ' Reset releases a lock
            Open f For Binary Access Read Lock Read Write As #1
            Reset
            Open f For Binary Access Read Lock Read Write As #1
            v = "reopened"
        Case 12  ' Reset and the Dir sequence
            v = Dir(Dr & "\*.txt")
            Reset
            v = v & "," & Dir()
        Case 13  ' FreeFile after Reset
            Open f For Input Shared As #1
            Open f For Input Shared As #2
            Reset
            v = FreeFile
        Case 14  ' Reset with files in all four modes
            Open Dr & "\o.txt" For Output As #1
            Open Dr & "\a.txt" For Append As #2
            Open f For Binary As #3
            Open Dr & "\r.dat" For Random As #4 Len = 4
            s = OpenCount() & ","
            Reset
            v = s & OpenCount()
    End Select
    s = "Reset." & k & "=" & Sh(v)
    Wipe
    T = s
    Exit Function
fail:
    e = Err.Number
    Wipe
    T = "Reset." & k & "=ERR" & e
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
