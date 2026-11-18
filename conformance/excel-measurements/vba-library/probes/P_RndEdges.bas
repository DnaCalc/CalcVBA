Attribute VB_Name = "P_RndEdges"
Option Explicit
' Additional RNG coercion edges. UNMEASURED until captured with Excel.
Private Const N As Long = 8
Private Function T(ByVal k As Long) As String
    Dim v As Variant, w As Variant
    On Error GoTo fail
    w = Rnd(-1)
    Select Case k
        Case 1: v = Sh(Rnd(Empty)) & "|" & Sh(Rnd(0))
        Case 2: v = Sh(Rnd(False)) & "|" & Sh(Rnd(0))
        Case 3: v = Sh(Rnd(-1E-300)) & "|" & Sh(Rnd(0))
        Case 4: v = Sh(Rnd(1E-300)) & "|" & Sh(Rnd(0))
        Case 5: v = Sh(Rnd(CDec(-1))) & "|" & Sh(Rnd(0))
        Case 6: Randomize Empty: v = Sh(Rnd())
        Case 7: Randomize CDec(255): v = Sh(Rnd())
        Case 8: Randomize "255": v = Sh(Rnd())
    End Select
    T = "RndEdges." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "RndEdges." & k & "=ERR" & Err.Number
End Function
Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
