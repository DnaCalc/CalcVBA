Attribute VB_Name = "P_Seek"
Option Explicit
' Seek function and Seek # statement (FileSystem; MS-VBAL 6.1.2.8.x). Files only under %TEMP%\cvp-lib\Seek.
Private Const N As Long = 20

Private Function Dr() As String
    Dr = Environ("TEMP") & "\cvp-lib\Seek"
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
    Dim v As Variant, s As String, e As Long, n As Integer, f As String
    Dim x As Long, y As Long, l As String, b As Byte, t2 As String, p As Long
    On Error GoTo fail
    Wipe
    Mk
    f = Dr & "\f.dat"
    n = FreeFile
    Select Case k
        Case 1: Open f For Binary As #n: v = Seek(n)                        ' Binary starts at 1
        Case 2   ' Binary advances with Put
            Open f For Binary As #n
            t2 = "abcde": Put #n, , t2
            v = Seek(n)
        Case 3   ' Seek statement then read
            Wr f, "0123456789"
            Open f For Binary As #n
            Seek #n, 4
            s = Seek(n) & ","
            Get #n, , b
            v = s & Chr$(b) & "," & Seek(n)
        Case 4   ' Random: next record number
            Open f For Random As #n Len = 4
            s = Seek(n) & ","
            Get #n, 2, y: s = s & Seek(n) & ","
            x = 1: Put #n, 5, x: s = s & Seek(n)
            v = s
        Case 5   ' Input mode after one line plus CRLF
            Wr f, "ab" & vbCrLf & "cd"
            Open f For Input As #n
            s = Seek(n) & ","
            Line Input #n, l
            v = s & Seek(n)
        Case 6   ' Output mode after Print
            Open f For Output As #n
            s = Seek(n) & ","
            Print #n, "abc"
            v = s & Seek(n)
        Case 7   ' Append mode starts at the end
            Wr f, "abcd"
            Open f For Append As #n
            s = Seek(n) & ","
            Print #n, "xy";
            v = s & Seek(n)
        Case 8: Wr f, "abc": Open f For Binary As #n: Seek #n, 0            ' position 0: error 5
        Case 9: Wr f, "abc": Open f For Binary As #n: Seek #n, -1           ' negative position: error 5
        Case 10  ' Seek past the end then Put grows the file
            Wr f, "abcde"
            Open f For Binary As #n
            Seek #n, 100
            t2 = "x": Put #n, , t2
            v = LOF(n) & "," & Seek(n)
        Case 11  ' Input mode: Seek past the end
            Wr f, "abc"
            Open f For Input As #n
            Seek #n, 10
            v = EOF(n) & "," & Seek(n)
        Case 12  ' Random: Seek statement picks the record
            Open f For Random As #n Len = 4
            x = 11: Put #n, 1, x
            x = 22: Put #n, 2, x
            x = 33: Put #n, 3, x
            Seek #n, 3
            s = Seek(n) & ","
            Get #n, , y
            v = s & y & "," & Seek(n)
        Case 13: v = Seek(5)                                                ' never opened: error 52
        Case 14: Open f For Binary As #n: Close #n: v = Seek(n)             ' after Close: error 52
        Case 15: Open f For Binary As #n: v = TypeName(Seek(n))             ' Long
        Case 16: Open f For Binary As #n: v = Seek(FileNumber:=n)           ' named argument
        Case 17  ' Output mode: Seek back and overwrite
            Open f For Output As #n
            Print #n, "abcdef";
            Seek #n, 3
            Print #n, "XY";
            Close #n
            v = Sh(Rd(f))
        Case 18: Wr f, "abc": Open f For Binary As #n: Seek #n, Null       ' Null position: error 94
        Case 19  ' Binary: Seek to 1 after reading rereads
            Wr f, "abc"
            Open f For Binary As #n
            Get #n, , b: Get #n, , b
            p = Seek(n)
            Seek #n, 1
            Get #n, , b
            v = p & "," & Chr$(b) & "," & Seek(n)
        Case 20  ' Random: Seek 0 is an error
            Open f For Random As #n Len = 4
            Seek #n, 0
    End Select
    s = "Seek." & k & "=" & Sh(v)
    Wipe
    T = s
    Exit Function
fail:
    e = Err.Number
    Wipe
    T = "Seek." & k & "=ERR" & e
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
