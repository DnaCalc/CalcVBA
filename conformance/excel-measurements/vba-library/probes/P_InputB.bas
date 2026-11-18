Attribute VB_Name = "P_InputB"
Option Explicit
' InputB / InputB$ (hidden module of VBA; file function). Files only under %TEMP%\cvp-lib\InputB.
Private Const N As Long = 18

Private Function Dr() As String
    Dr = Environ("TEMP") & "\cvp-lib\InputB"
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
    Dim v As Variant, s As String, e As Long, n As Integer, f As String, b() As Byte
    On Error GoTo fail
    Wipe
    Mk
    f = Dr & "\f.txt"
    n = FreeFile
    Select Case k
        Case 1   ' four bytes: two characters, first is "ab" little-endian
            Wr f, "abcd": Open f For Binary As #n
            s = InputB(4, #n)
            v = Len(s) & "," & LenB(s) & "," & Sh(s)
        Case 2   ' $ form
            Wr f, "abcd": Open f For Binary As #n
            s = InputB$(4, #n)
            v = Len(s) & "," & LenB(s) & "," & Sh(s)
        Case 3   ' one byte
            Wr f, "abcd": Open f For Binary As #n
            s = InputB(1, #n)
            v = Len(s) & "," & LenB(s)
        Case 4   ' Input versus InputB on the same bytes
            Wr f, "abcd": Open f For Input As #n
            v = Len(Input(2, #n)) & "," & LenB(InputB(2, #n))
        Case 5: Wr f, "ab": Open f For Input As #n: v = InputB(3, #n)             ' past EOF: error 62
        Case 6: Wr f, "ab": Open f For Input As #n: s = InputB(2, #n): v = InputB(1, #n)   ' at EOF: error 62
        Case 7: Open f For Output As #n: v = InputB(1, #n)                         ' Output mode: error 54
        Case 8: Wr f, "abc": Open f For Input As #n: v = InputB(-1, #n)           ' negative count: error 5
        Case 9: Wr f, "abc": Open f For Input As #n: v = InputB(Null, #n)         ' Null count: error 94
        Case 10: v = InputB(1, #5)                                                 ' not open: error 52
        Case 11: Wr f, "abcd": Open f For Input As #n: v = Sh(InputB(0, #n))       ' zero count
        Case 12   ' odd counts in sequence
            Wr f, "abcd": Open f For Binary As #n
            s = InputB(3, #n)
            v = LenB(s) & "," & LenB(InputB(1, #n))
        Case 13   ' ANSI view with StrConv
            Wr f, "abcd": Open f For Binary As #n
            v = Sh(StrConv(InputB(4, #n), vbUnicode))
        Case 14   ' result types
            Wr f, "abcd": Open f For Binary As #n
            v = TypeName(InputB(2, #n)) & "," & TypeName(InputB$(2, #n))
        Case 15   ' assigned to a byte array
            Wr f, "abcd": Open f For Binary As #n
            b = InputB$(4, #n)
            v = UBound(b) & ":" & b(0) & "," & b(3)
        Case 16: Wr f, "abc": Open f For Input As #n: Close #n: v = InputB(1, #n)  ' after Close: error 52
        Case 17   ' NUL and high bytes survive
            Wr f, "a" & Chr$(0) & Chr$(255) & "b": Open f For Binary As #n
            b = InputB$(4, #n)
            v = b(0) & "," & b(1) & "," & b(2) & "," & b(3)
        Case 18   ' InputB moves Seek by bytes
            Wr f, "abcdef": Open f For Binary As #n
            s = InputB(4, #n)
            v = Seek(n) & "," & EOF(n)
    End Select
    s = "InputB." & k & "=" & Sh(v)
    Wipe
    T = s
    Exit Function
fail:
    e = Err.Number
    Wipe
    T = "InputB." & k & "=ERR" & e
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
