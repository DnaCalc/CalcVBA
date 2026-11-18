Attribute VB_Name = "P_Timer"
Option Explicit
' Timer (VBA.DateTime; MS-VBAL 6.1.2.x).
Private Const N As Long = 12

Private Function T(ByVal k As Long) As String
    Dim v As Variant, a As Variant, b As Variant, c As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = TypeName(Timer)   ' type is Single
        Case 2: v = VarType(Timer) = vbSingle   ' VarType
        Case 3: v = Timer >= 0 And Timer < 86400   ' seconds since midnight range
        Case 4: v = Timer >= 0   ' non-negative
        Case 5: a = Timer: v = (a <= Timer)   ' two reads are non-decreasing (fails only across midnight)
        Case 6: v = Abs(Timer - (CDbl(Time) * 86400)) < 2   ' agrees with Time within two seconds (fails only across midnight)
        Case 7: v = TypeName(Timer + 1)   ' Single + Integer stays Single
        Case 8: v = TypeName(Timer * 2#)   ' Single * Double becomes Double
        Case 9: v = TypeName(VBA.Timer)   ' qualified call
        Case 10: v = Int(Timer) = Fix(Timer)   ' Int and Fix agree for a non-negative value
        Case 11: v = Timer = CSng(Timer)   ' value is already Single precision (fails only if the clock ticks between reads)
        Case 12: v = Timer - Fix(Timer) < 1   ' fractional part below one
    End Select
    T = "Timer." & k & "=" & Sh(v)
    Exit Function
fail:
    Dim en As Long
    en = Err.Number
    T = "Timer." & k & "=ERR" & en
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
