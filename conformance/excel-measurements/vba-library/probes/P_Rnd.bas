Attribute VB_Name = "P_Rnd"
Option Explicit
' Rnd (VBA.Math). Deterministic seeds only.
Private Const N As Long = 35

Private Function Seq(ByVal n As Long) As String
    Dim i As Long
    For i = 1 To n
        Seq = Seq & "," & Sh(Rnd())
    Next
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    Dim i As Long, first As Single, w As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Sh(Rnd(-1)) & Seq(3)                              ' seed -1, then 3 values
        Case 2: v = Sh(Rnd(-2)) & Seq(3)                              ' seed -2
        Case 3: v = Sh(Rnd(-1.1)) & Seq(3)                            ' seed -1.1
        Case 4: v = Sh(Rnd(-1E+10)) & Seq(3)                          ' seed -1E10
        Case 5: v = Sh(Rnd(-0.5)) & Seq(3)                            ' seed -0.5
        Case 6: v = Sh(Rnd(-100000)) & Seq(3)                         ' seed -100000
        Case 7: v = Sh(Rnd(-1)) & Sh(Rnd(0)) & Sh(Rnd(0)) & Sh(Rnd(0))' Rnd(0) repeats the last value
        Case 8: v = Sh(Rnd(-1)) & Sh(Rnd(1)) & Sh(Rnd(2)) & Sh(Rnd(1000))' positive arguments act as Rnd()
        Case 9: v = Sh(Rnd(-1)) & Sh(Rnd()) & Sh(Rnd(0)) & Sh(Rnd())  ' Rnd(0) after a value
        Case 10: v = Sh(Rnd(-1)) & Sh(Rnd(-1)) & Sh(Rnd(-1))          ' same negative seed gives the same value
        Case 11: v = Sh(Rnd(-1)) & Sh(Rnd(-2)) & Seq(2)               ' seed change in sequence
        Case 12: Rnd -1: Randomize 5: v = Seq(3)                      ' Randomize 5 after Rnd -1
        Case 13: Rnd -1: Randomize 0: v = Seq(3)                      ' Randomize 0
        Case 14: Rnd -1: Randomize -1: v = Seq(3)                     ' Randomize -1
        Case 15: Rnd -1: Randomize 1.5: v = Seq(3)                    ' Randomize 1.5
        Case 16: Rnd -1: Randomize 5: Randomize 5: v = Seq(3)         ' Randomize 5 twice (state-dependent)
        Case 17: Rnd -1: Randomize 12345.678: v = Seq(3)              ' Randomize with a fraction
        Case 18: Rnd -1: Randomize 1E+10: v = Seq(3)                  ' Randomize 1E10
        Case 19: Rnd -1: Randomize -1E+10: v = Seq(3)                 ' Randomize -1E10
        Case 20: Rnd -1: Randomize "7": v = Seq(3)                    ' Randomize numeric string
        Case 21: Rnd -1: Randomize CSng(5): v = Seq(3)                ' Randomize Single 5
        Case 22: Rnd -1: Randomize CCur(5): v = Seq(3)                ' Randomize Currency 5
        Case 23: Rnd -1: Randomize 5&: v = Seq(3)                     ' Randomize Long 5
        Case 24: Rnd -1: Randomize 65536: v = Seq(3)                  ' Randomize 65536
        Case 25: Rnd -1: Randomize Number:=5: v = Seq(3)              ' Randomize named argument
        Case 26: Randomize Null: v = "no error"                       ' Randomize Null: error 94
        Case 27: Randomize "abc": v = "no error"                      ' Randomize non-numeric: error 13
        Case 28: Rnd -1: Randomize Empty: v = Seq(3)                  ' Randomize Empty
        Case 29: v = Rnd(Null)                                        ' Rnd Null: error 94
        Case 30: v = Rnd("abc")                                       ' Rnd non-numeric: error 13
        Case 31: v = Sh(Rnd(-1)) & Sh(Rnd("-1")) & Sh(Rnd(-1&)) & Sh(Rnd(CCur(-1)))' Rnd(-1) in several types
        Case 32: v = Sh(Rnd(Number:=-1)) & Seq(2)                     ' named argument
        Case 33: Rnd -1: For i = 1 To 1000: w = Rnd(): Next: v = Sh(w)' 1000th value
        Case 34: first = Rnd(-1): For i = 1 To 16777215: w = Rnd(): Next: v = Sh(first) & "|" & Sh(Rnd())' LCG period 2^24: returns to first
        Case 35: Rnd -1: v = TypeName(Rnd())                          ' result type Single
    End Select
    T = "Rnd." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Rnd." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
