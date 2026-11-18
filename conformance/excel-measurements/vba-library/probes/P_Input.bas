Attribute VB_Name = "P_Input"
Option Explicit
' Input / Input$ (hidden module of VBA; file function). Files only under %TEMP%\cvp-lib\Input.
Private Const N As Long = 22

Private Function Dr() As String
    Dr = Environ("TEMP") & "\cvp-lib\Input"
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
    Dim v As Variant, s As String, e As Long, n As Integer, f As String, l As String
    On Error GoTo fail
    Wipe
    Mk
    f = Dr & "\f.txt"
    n = FreeFile
    Select Case k
        Case 1: Wr f, "abcdef": Open f For Input As #n: v = Input(3, #n)           ' typical
        Case 2: Wr f, "abcdef": Open f For Input As #n: v = Input$(3, #n)          ' $ form
        Case 3: Wr f, "abcdef": Open f For Input As #n: v = TypeName(Input(3, #n)) & "," & TypeName(Input$(3, #n))   ' result types
        Case 4: Wr f, "abc": Open f For Input As #n: v = Sh(Input(0, #n))          ' zero count
        Case 5: Wr f, "ab": Open f For Input As #n: v = Input(3, #n)               ' past EOF: error 62
        Case 6: Wr f, "ab": Open f For Input As #n: s = Input(2, #n): v = Input(1, #n)   ' at EOF: error 62
        Case 7: Wr f, "abcdef": Open f For Binary As #n: v = Input(4, #n)          ' Binary mode
        Case 8: Open f For Output As #n: v = Input(1, #n)                          ' Output mode: error 54
        Case 9: Wr f, "abcdefgh": Open f For Random As #n Len = 4: v = Input(3, #n)   ' Random mode
        Case 10: Wr f, "a" & vbCrLf & "b": Open f For Input As #n: v = Sh(Input(4, #n))   ' CRLF is returned as is
        Case 11: Wr f, "ab" & Chr$(26) & "cd": Open f For Input As #n: v = Sh(Input(5, #n))   ' Ctrl-Z in Input mode
        Case 12: Wr f, "abc": Open f For Input As #n: v = Input(-1, #n)            ' negative count: error 5
        Case 13: Wr f, "abc": Open f For Input As #n: v = Input(Null, #n)          ' Null count: error 94
        Case 14: v = Input(1, #5)                                                  ' not open: error 52
        Case 15: Wr f, String$(1000, "q"): Open f For Input As #n: v = Len(Input(1000, #n)) & "," & (Input(0, #n) = "")   ' whole 1000-byte file
        Case 16: Wr f, "a" & Chr$(0) & "b": Open f For Binary As #n: v = Sh(Input(3, #n))   ' NUL byte
        Case 17: Wr f, "a" & Chr$(200) & "b": Open f For Binary As #n: s = Input(3, #n): v = Len(s) & ":" & AscW(Mid$(s, 2, 1))   ' high ANSI byte (locale)
        Case 18: Wr f, "abcdef": Open f For Input As #n: v = Input(2, n)           ' file number without #
        Case 19   ' reading in pieces and then a line
            Wr f, "abcd" & vbCrLf & "efgh"
            Open f For Input As #n
            s = Input(2, #n)
            Line Input #n, l
            v = s & "|" & l & "|" & Input(2, #n)
        Case 20: Wr f, "abc": Open f For Input As #n: Close #n: v = Input(1, #n)   ' after Close: error 52
        Case 21: Wr f, "abc": Open f For Input As #n: v = Input("2", #n)           ' numeric string count
        Case 22   ' Input moves Loc and Seek
            Wr f, "abcdef"
            Open f For Input As #n
            s = Input(2, #n)
            v = EOF(n) & "," & Seek(n)
    End Select
    s = "Input." & k & "=" & Sh(v)
    Wipe
    T = s
    Exit Function
fail:
    e = Err.Number
    Wipe
    T = "Input." & k & "=ERR" & e
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
