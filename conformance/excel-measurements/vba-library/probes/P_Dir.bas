Attribute VB_Name = "P_Dir"
Option Explicit
' Dir (FileSystem; MS-VBAL 6.1.2.8.x). Files only under %TEMP%\cvp-lib\Dir.
Private Const N As Long = 22

Private Function Dr() As String
    Dr = Environ("TEMP") & "\cvp-lib\Dir"
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

' a.txt b.txt ab1.txt c.dat (normal), ro.txt (read-only), h.txt (hidden), sys.txt (system), sub\ (folder) with sub\in.txt
Private Sub Setup()
    Mk
    Wr Dr & "\a.txt", "a"
    Wr Dr & "\b.txt", "b"
    Wr Dr & "\ab1.txt", "ab1"
    Wr Dr & "\c.dat", "c"
    Wr Dr & "\ro.txt", "ro"
    Wr Dr & "\h.txt", "h"
    Wr Dr & "\sys.txt", "sys"
    MkDir Dr & "\sub"
    Wr Dr & "\sub\in.txt", "in"
    SetAttr Dr & "\ro.txt", vbReadOnly
    SetAttr Dr & "\h.txt", vbHidden
    SetAttr Dr & "\sys.txt", vbSystem
End Sub

' All names Dir returns for a pattern, in the order returned; attr = -1 omits the argument.
Private Function Lst(ByVal pat As String, ByVal attr As Long) As String
    Dim f As String, s As String, c As Long
    If attr = -1 Then f = Dir(pat) Else f = Dir(pat, attr)
    Do While Len(f) > 0 And c < 100
        s = s & f & "|"
        c = c + 1
        f = Dir()
    Loop
    Lst = c & ":" & s
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String, e As Long, v2 As Variant, v3 As Variant
    On Error GoTo fail
    Wipe
    Setup
    Select Case k
        Case 1: v = Lst(Dr & "\*.txt", -1)                                  ' default attributes: normal files only
        Case 2: v = Lst(Dr & "\*", vbNormal)                                ' explicit vbNormal
        Case 3: v = Lst(Dr & "\*", vbHidden)                                ' hidden added
        Case 4: v = Lst(Dr & "\*", vbReadOnly)                              ' read-only added
        Case 5: v = Lst(Dr & "\*", vbSystem)                                ' system added
        Case 6: v = Lst(Dr & "\*", vbDirectory)                             ' folders added
        Case 7: v = Lst(Dr & "\*", vbHidden + vbSystem + vbReadOnly + vbDirectory)   ' everything
        Case 8: v = Lst(Dr & "\?.txt", -1)                                  ' ? matches one character
        Case 9: v = Lst(Dr & "\a*", -1)                                     ' prefix wildcard
        Case 10: v = Lst(Dr & "\", -1)                                      ' folder with trailing backslash
        Case 11: v = Dir(Dr) & "/" & Dir(Dr, vbDirectory)                   ' folder without backslash: only with vbDirectory
        Case 12
            v = Dir(Dr & "\a.txt"): v2 = Dir()                              ' sequence ends with ""
            On Error Resume Next
            v3 = Dir()
            e = Err.Number
            On Error GoTo fail
            v = v & "/" & v2 & "/" & e                                      ' Dir() after the end: error 5
        Case 13: v = Dir()                                                  ' no prior pattern: error 5
        Case 14: v = Dir(Dr & "\zzz*"): v = TypeName(v) & ":" & Len(v)     ' no match gives ""
        Case 15: v = Dir(Dr & "\nodir\*")                                   ' missing folder
        Case 16: v = Dir(Dr & "\sub\in.txt")                                ' name only, no path
        Case 17: v = Dir(Dr & "\A.TXT")                                     ' name as stored on disk
        Case 18: v = Lst(Dr & "\*.txt", vbHidden)                           ' hidden filter on a pattern
        Case 19: v = Dir(PathName:=Dr & "\c.dat", Attributes:=vbNormal)     ' named arguments
        Case 20: v = Dir(Null)                                              ' Null: error 94
        Case 21: v = Dir("") & "/" & TypeName(Dir(""))                      ' empty pattern
        Case 22
            v = Dir(Dr & "\*.txt"): v2 = Dir(Dr & "\*.dat"): v3 = Dir()    ' a new pattern restarts the sequence
            v = v & "/" & v2 & "/" & v3
    End Select
    s = "Dir." & k & "=" & Sh(v)
    Wipe
    T = s
    Exit Function
fail:
    e = Err.Number
    Wipe
    T = "Dir." & k & "=ERR" & e
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
