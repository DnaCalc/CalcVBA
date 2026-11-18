Attribute VB_Name = "P_FreeFile"
Option Explicit
' FreeFile (FileSystem; MS-VBAL 6.1.2.8.x). Files only under %TEMP%\cvp-lib\FreeFile.
Private Const N As Long = 20

Private Function Dr() As String
    Dr = Environ("TEMP") & "\cvp-lib\FreeFile"
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
    Dim v As Variant, s As String, e As Long, n As Integer, m As Integer, f As String, i As Long
    On Error GoTo fail
    Wipe
    Mk
    f = Dr & "\f.txt"
    Wr f, "x"
    Select Case k
        Case 1: v = FreeFile                                          ' nothing open: 1
        Case 2: v = FreeFile(0) & "," & FreeFile(1)                   ' ranges: 1 and 256
        Case 3: v = (FreeFile = FreeFile)                             ' FreeFile does not reserve
        Case 4: Open f For Input As #1: v = FreeFile & "," & FreeFile(1)   ' #1 taken: 2 and still 256
        Case 5   ' the upper range counts from 256
            n = FreeFile(1)
            Open f For Input As #n
            v = n & "," & FreeFile(1) & "," & FreeFile(0)
        Case 6   ' lowest free number is reused
            Open f For Input As #1
            Open f For Input Shared As #2
            Close #1
            v = FreeFile
        Case 7: v = FreeFile(2)                                       ' range other than 0 or 1: error 5
        Case 8: v = FreeFile(-1)                                      ' negative range: error 5
        Case 9: v = FreeFile(Null)                                    ' Null: error 94
        Case 10: v = FreeFile(Empty)                                  ' Empty counts as 0
        Case 11: v = FreeFile("1")                                    ' numeric string
        Case 12: v = FreeFile(1.5)                                    ' rounds to 2: error 5
        Case 13: v = FreeFile(0.4)                                    ' rounds to 0
        Case 14: v = TypeName(FreeFile) & "," & TypeName(FreeFile(1)) ' Integer
        Case 15   ' three open: next is 4
            Open f For Input Shared As #1
            Open f For Input Shared As #2
            Open f For Input Shared As #3
            v = FreeFile
        Case 16   ' a gap: #1 and #3 open, 2 is free
            Open f For Input Shared As #1
            Open f For Input Shared As #3
            v = FreeFile
        Case 17: v = FreeFile(RangeNumber:=1)                         ' named argument
        Case 18   ' all 255 low numbers taken: error 67
            For i = 1 To 255
                Open f For Input Shared As #i
            Next
            v = FreeFile
        Case 19   ' all 255 low numbers taken: the upper range is still free
            For i = 1 To 255
                Open f For Input Shared As #i
            Next
            v = FreeFile(1)
        Case 20   ' after Close of everything the numbers come back
            Open f For Input Shared As #1
            Open f For Input Shared As #2
            Close
            v = FreeFile
    End Select
    s = "FreeFile." & k & "=" & Sh(v)
    Wipe
    T = s
    Exit Function
fail:
    e = Err.Number
    Wipe
    T = "FreeFile." & k & "=ERR" & e
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
