Attribute VB_Name = "P_Time"
Option Explicit
' Time / Time$ (VBA.DateTime; MS-VBAL 6.1.2.x).
Private Const N As Long = 14

Private Function T(ByVal k As Long) As String
    Dim v As Variant, a As Variant, b As Variant, c As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = TypeName(Time)   ' type is Date
        Case 2: v = VarType(Time) = vbDate   ' VarType
        Case 3: v = TypeName(Time$)   ' Time$ returns String
        Case 4: v = Len(Time$)   ' Time$ is hh:mm:ss, 8 chars
        Case 5: v = Mid$(Time$, 3, 1) & Mid$(Time$, 6, 1)   ' separators of Time$ are colons
        Case 6: v = CDbl(Time) >= 0 And CDbl(Time) < 1   ' no date part
        Case 7: v = Year(Time)   ' date part is serial 0, 1899
        Case 8: v = Day(Time)   ' day of serial 0
        Case 9: v = Abs(CDbl(Time) * 86400 - Fix(CDbl(Time) * 86400 + 0.5)) < 0.001   ' whole-second resolution
        Case 10: v = Hour(Time) >= 0 And Hour(Time) <= 23   ' Hour range
        Case 11: v = Val(Left$(Time$, 2)) >= 0 And Val(Left$(Time$, 2)) <= 23   ' hours of Time$
        Case 12: v = IsDate(Time$)   ' IsDate on Time$
        Case 13: v = Int(Now) + Time >= Int(Now)   ' adds back to a date
        Case 14: v = TypeName(VBA.Time)   ' qualified call
    End Select
    T = "Time." & k & "=" & Sh(v)
    Exit Function
fail:
    Dim en As Long
    en = Err.Number
    T = "Time." & k & "=ERR" & en
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
