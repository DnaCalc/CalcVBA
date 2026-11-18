Attribute VB_Name = "P_Now"
Option Explicit
' Now (VBA.DateTime; MS-VBAL 6.1.2.x).
Private Const N As Long = 12

Private Function T(ByVal k As Long) As String
    Dim v As Variant, a As Variant, b As Variant, c As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = TypeName(Now)   ' type is Date
        Case 2: v = VarType(Now) = vbDate   ' VarType
        Case 3: v = Int(Now) = Date   ' date part equals Date (fails only across midnight)
        Case 4: v = Now >= Date   ' not before today's midnight
        Case 5: v = Now - Date < 1   ' less than a day past midnight
        Case 6: v = CDbl(Now) > 43831   ' later than 2020-01-01
        Case 7: v = Abs(CDbl(Now) * 86400 - Fix(CDbl(Now) * 86400 + 0.5)) < 0.001   ' whole-second resolution
        Case 8: v = Abs((CDbl(Now) - Int(CDbl(Now))) - CDbl(Time)) < 0.0001   ' time part equals Time (fails only across a second tick)
        Case 9: v = Year(Now) = Year(Date)   ' Year of Now
        Case 10: a = Now: v = (a <= Now)   ' two reads are non-decreasing
        Case 11: v = IsDate(Now)   ' IsDate
        Case 12: v = TypeName(Now + 1)   ' Date + Double stays Date
    End Select
    T = "Now." & k & "=" & Sh(v)
    Exit Function
fail:
    Dim en As Long
    en = Err.Number
    T = "Now." & k & "=ERR" & en
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
