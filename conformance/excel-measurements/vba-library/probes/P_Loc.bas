Attribute VB_Name = "P_Loc"
Option Explicit
' Loc (FileSystem; MS-VBAL 6.1.2.8.x). Files only under %TEMP%\cvp-lib\Loc.
Private Const N As Long = 18

Private Type Rec
    id As Long
    nm As String * 6
End Type

Private Function Dr() As String
    Dr = Environ("TEMP") & "\cvp-lib\Loc"
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
    Dim x As Long, y As Long, l As String, b As Byte, r As Rec, t2 As String
    On Error GoTo fail
    Wipe
    Mk
    f = Dr & "\f.dat"
    n = FreeFile
    Select Case k
        Case 1   ' Binary: Loc is the last byte read or written
            Open f For Binary As #n
            s = Loc(n) & ","
            t2 = "abcde": Put #n, , t2: s = s & Loc(n) & ","
            Seek #n, 1
            Get #n, , b: s = s & Loc(n) & ","
            Get #n, , b: s = s & Loc(n)
            v = s
        Case 2   ' Random: Loc is the last record number
            Open f For Random As #n Len = 4
            s = Loc(n) & ","
            x = 7: Put #n, 3, x: s = s & Loc(n) & ","
            Get #n, 1, y: s = s & Loc(n) & ","
            Get #n, , y: s = s & Loc(n)
            v = s
        Case 3   ' Input mode: position divided by 128
            Wr f, String$(300, "a") & vbCrLf & String$(200, "b")
            Open f For Input As #n
            s = Loc(n) & ","
            Line Input #n, l: s = s & Loc(n) & ","
            Line Input #n, l: s = s & Loc(n)
            v = s
        Case 4   ' Output mode after Print of 200 characters
            Open f For Output As #n
            Print #n, String$(200, "a")
            s = Loc(n) & ","
            Print #n, String$(200, "a")
            s = s & Loc(n)
            v = s
        Case 5   ' Append mode after existing 300 bytes
            Wr f, String$(300, "a")
            Open f For Append As #n
            s = Loc(n) & ","
            Print #n, "xyz"
            s = s & Loc(n)
            v = s
        Case 6: v = Loc(5)                                                ' never opened: error 52
        Case 7: Open f For Binary As #n: Close #n: v = Loc(n)             ' after Close: error 52
        Case 8: Open f For Binary As #n: v = TypeName(Loc(n))             ' Long
        Case 9   ' Binary: Seek then Loc is one less
            Wr f, "0123456789"
            Open f For Binary As #n
            Seek #n, 6
            v = Loc(n)
        Case 10  ' Random with records: sequential Put counts records
            Open f For Random As #n Len = Len(r)
            r.id = 1: r.nm = "one": Put #n, , r: s = Loc(n) & ","
            r.id = 2: r.nm = "two": Put #n, , r: s = s & Loc(n) & ","
            Get #n, 1, r: s = s & Loc(n) & ":" & r.id
            v = s
        Case 11: Wr f, "x": Open f For Input As #n: v = Loc(FileNumber:=n)    ' named argument
        Case 12: v = Loc(Null)                                            ' Null: error 94
        Case 13  ' Input mode at the end of a 300 byte file
            Wr f, String$(300, "a")
            Open f For Input As #n
            Line Input #n, l
            v = Loc(n)
        Case 14: Wr f, "": Open f For Input As #n: v = Loc(n)             ' empty file
        Case 15  ' Binary Get of a string
            Wr f, "abcdefgh"
            Open f For Binary As #n
            t2 = Space$(3)
            Get #n, , t2
            v = Loc(n) & ":" & t2
        Case 16: v = Loc(0)                                               ' file number 0: error 52
        Case 17  ' Random: Get with record number beyond the end
            Wr f, "abcdefgh"
            Open f For Random As #n Len = 4
            Get #n, 2, y
            s = Loc(n) & ","
            Get #n, 3, y
            v = s & Loc(n)
        Case 18  ' Binary Put of a Long at position 10
            Open f For Binary As #n
            x = 5
            Put #n, 10, x
            v = Loc(n)
    End Select
    s = "Loc." & k & "=" & Sh(v)
    Wipe
    T = s
    Exit Function
fail:
    e = Err.Number
    Wipe
    T = "Loc." & k & "=ERR" & e
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
