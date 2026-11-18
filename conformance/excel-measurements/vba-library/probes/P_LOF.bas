Attribute VB_Name = "P_LOF"
Option Explicit
' LOF (FileSystem; MS-VBAL 6.1.2.8.x). Files only under %TEMP%\cvp-lib\LOF.
Private Const N As Long = 18

Private Function Dr() As String
    Dr = Environ("TEMP") & "\cvp-lib\LOF"
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
    Dim x As Long, t2 As String
    On Error GoTo fail
    Wipe
    Mk
    f = Dr & "\f.dat"
    n = FreeFile
    Select Case k
        Case 1   ' Binary after writing five bytes
            Open f For Binary As #n
            t2 = "abcde": Put #n, , t2
            v = LOF(n)
        Case 2: Wr f, "0123456789": Open f For Input As #n: v = LOF(n)     ' Input mode
        Case 3   ' Output mode after a Print
            Open f For Output As #n
            Print #n, "abc"
            v = LOF(n)
        Case 4   ' Append mode: existing bytes plus a Print
            Wr f, "0123456789"
            Open f For Append As #n
            s = LOF(n) & ","
            Print #n, "ab";
            v = s & LOF(n)
        Case 5   ' Random: Put of record 3 with Len 4
            Open f For Random As #n Len = 4
            x = 1: Put #n, 3, x
            v = LOF(n)
        Case 6: Wr f, "": Open f For Input As #n: v = LOF(n)               ' empty file
        Case 7: v = LOF(7)                                                  ' never opened: error 52
        Case 8: Open f For Binary As #n: v = TypeName(LOF(n))               ' Long
        Case 9: Open f For Binary As #n: Close #n: v = LOF(n)               ' after Close: error 52
        Case 10  ' equals FileLen after Close
            Open f For Binary As #n
            t2 = "abcdefg": Put #n, , t2
            Close #n
            v = (FileLen(f) = 7)
        Case 11  ' Seek far past the end then Put one byte
            Open f For Binary As #n
            Seek #n, 11
            t2 = "x": Put #n, , t2
            v = LOF(n)
        Case 12: Wr f, "abc": Open f For Input As #n: v = LOF(FileNumber:=n)   ' named argument
        Case 13: v = LOF(Null)                                              ' Null: error 94
        Case 14  ' Output truncates an existing file
            Wr f, String$(100, "x")
            Open f For Output As #n
            v = LOF(n)
        Case 15  ' Random Len 10 over a 25-byte file
            Wr f, String$(25, "x")
            Open f For Random As #n Len = 10
            v = LOF(n)
        Case 16: v = LOF(0)                                                 ' file number 0: error 52
        Case 17  ' Binary Put grows the file
            Wr f, "abc"
            Open f For Binary As #n
            s = LOF(n) & ","
            Seek #n, 4
            t2 = "defgh": Put #n, , t2
            v = s & LOF(n)
        Case 18  ' a 40000-byte file: Long result
            Wr f, String$(40000, "x")
            Open f For Binary As #n
            v = LOF(n) & ":" & TypeName(LOF(n))
    End Select
    s = "LOF." & k & "=" & Sh(v)
    Wipe
    T = s
    Exit Function
fail:
    e = Err.Number
    Wipe
    T = "LOF." & k & "=ERR" & e
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
