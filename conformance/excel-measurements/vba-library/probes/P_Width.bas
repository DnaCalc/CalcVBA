Attribute VB_Name = "P_Width"
Option Explicit
' Width # statement (hidden module of VBA). Files only under %TEMP%\cvp-lib\Width.
Private Const N As Long = 18

Private Function Dr() As String
    Dr = Environ("TEMP") & "\cvp-lib\Width"
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
    Dim v As Variant, s As String, e As Long, n As Integer, f As String, w As Integer
    On Error GoTo fail
    Wipe
    Mk
    f = Dr & "\f.txt"
    n = FreeFile
    Select Case k
        Case 1   ' width 10, 16 characters
            Open f For Output As #n
            Width #n, 10
            Print #n, "ABCDEFGHIJKLMNOP"
            Close #n
            v = Sh(Rd(f))
        Case 2   ' width 0: no wrapping, 300 characters
            Open f For Output As #n
            Width #n, 0
            Print #n, String$(300, "x")
            Close #n
            v = Len(Rd(f))
        Case 3   ' default width: 300 characters
            Open f For Output As #n
            Print #n, String$(300, "x")
            Close #n
            v = Len(Rd(f))
        Case 4   ' several items with semicolons
            Open f For Output As #n
            Width #n, 5
            Print #n, "ab"; "cd"; "ef"; "gh"
            Close #n
            v = Sh(Rd(f))
        Case 5   ' comma print zones with a small width
            Open f For Output As #n
            Width #n, 5
            Print #n, "a", "b"
            Close #n
            v = Sh(Rd(f))
        Case 6: Wr f, "abc": Open f For Input As #n: Width #n, 5                   ' Input mode
        Case 7: Open f For Output As #n: Width #n, 256                             ' above 255: error 5
        Case 8: Open f For Output As #n: Width #n, -1                              ' negative: error 5
        Case 9: Width #7, 10                                                       ' not open: error 52
        Case 10: Open f For Binary As #n: Width #n, 10                             ' Binary mode
        Case 11: Open f For Random As #n Len = 4: Width #n, 10                     ' Random mode
        Case 12   ' Append mode
            Wr f, "ab"
            Open f For Append As #n
            Width #n, 3
            Print #n, "WXYZ"
            Close #n
            v = Sh(Rd(f))
        Case 13   ' width changed in the middle
            Open f For Output As #n
            Width #n, 4
            Print #n, "aaaaaa";
            Width #n, 0
            Print #n, "bbbbbbbb"
            Close #n
            v = Sh(Rd(f))
        Case 14: Open f For Output As #n: Width #n, Null                           ' Null: error 94
        Case 15   ' Write # with a small width
            Open f For Output As #n
            Width #n, 5
            Write #n, "abcdefghij", "klm"
            Close #n
            v = Sh(Rd(f))
        Case 16   ' text of exactly the width
            Open f For Output As #n
            Width #n, 5
            Print #n, "12345"
            Print #n, "67"
            Close #n
            v = Sh(Rd(f))
        Case 17   ' width 1
            Open f For Output As #n
            Width #n, 1
            Print #n, "abc"
            Close #n
            v = Sh(Rd(f))
        Case 18   ' Width 255 and 400 characters
            Open f For Output As #n
            Width #n, 255
            Print #n, String$(400, "y")
            Close #n
            s = Rd(f)
            v = Len(s) & ":" & InStr(s, vbCrLf)
    End Select
    s = "Width." & k & "=" & Sh(v)
    Wipe
    T = s
    Exit Function
fail:
    e = Err.Number
    Wipe
    T = "Width." & k & "=ERR" & e
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
