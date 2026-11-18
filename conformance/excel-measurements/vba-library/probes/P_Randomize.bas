Attribute VB_Name = "P_Randomize"
Option Explicit
' Randomize (VBA.Math). Deterministic: Rnd -1 first, then Randomize n, then values.
Private Const N As Long = 22

Private Function Seq(ByVal n As Long) As String
    Dim i As Long
    For i = 1 To n
        Seq = Seq & "," & Sh(Rnd())
    Next
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: Rnd -1: Randomize 1: v = Seq(3)                       ' Randomize 1
        Case 2: Rnd -1: Randomize 2: v = Seq(3)                       ' Randomize 2
        Case 3: Rnd -1: Randomize 5: v = Seq(3)                       ' Randomize 5
        Case 4: Rnd -1: Randomize -5: v = Seq(3)                      ' Randomize -5
        Case 5: Rnd -1: Randomize 0.1: v = Seq(3)                     ' Randomize 0.1
        Case 6: Rnd -1: Randomize 255: v = Seq(3)                     ' Randomize 255
        Case 7: Rnd -1: Randomize 256: v = Seq(3)                     ' Randomize 256
        Case 8: Rnd -1: Randomize 65535: v = Seq(3)                   ' Randomize 65535
        Case 9: Rnd -1: Randomize 2147483647: v = Seq(3)              ' Randomize Long max
        Case 10: Rnd -1: Randomize 1E+300: v = Seq(3)                 ' Randomize huge
        Case 11: Rnd -1: Randomize 1E-300: v = Seq(3)                 ' Randomize tiny
        Case 12: Rnd -1: Randomize 3.14159265358979: v = Seq(3)       ' Randomize pi
        Case 13: Rnd -1: Randomize 5: v = Seq(1): Randomize 5: v = v & "|" & Seq(1)' repeat Randomize 5 (state-dependent)
        Case 14: Rnd -1: Randomize CSng(0.1): v = Seq(3)              ' Randomize Single 0.1
        Case 15: Rnd -1: Randomize CDate(5): v = Seq(3)               ' Randomize Date
        Case 16: Rnd -1: Randomize True: v = Seq(3)                   ' Randomize Boolean
        Case 17: Rnd -1: Randomize Empty: v = Seq(3)                  ' Randomize Empty
        Case 18: Rnd -1: Randomize Number:=7: v = Seq(3)              ' named argument
        Case 19: Randomize Null: v = "no error"                       ' Null: error 94
        Case 20: Randomize "abc": v = "no error"                      ' non-numeric string: error 13
        Case 21: Randomize: v = TypeName(Rnd())                       ' no argument (timer seed): only the type is stable
        Case 22: Randomize: v = (Rnd() >= 0 And Rnd() < 1)            ' no argument: value in range
    End Select
    T = "Randomize." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Randomize." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
