Attribute VB_Name = "P_RoundEdges"
Option Explicit
' Round signed zero, result types and maximum digits measured in Excel.
Private Const N As Long = 24

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Round(-0.05, 1)
        Case 2: v = Round(-0.005, 2)
        Case 3: v = Round(-1E-300, 0)
        Case 4: v = Round(-1E-300, 1)
        Case 5: v = Round(-0.49, 0)
        Case 6: v = Round(-0.049, 1)
        Case 7: v = Round(CSng(-0.05), 1)
        Case 8: v = Round(CSng(-1E-30), 1)
        Case 9: v = Round(CDate(-0.05), 1)
        Case 10: v = Round(CDate(2.55), 1)
        Case 11: v = Round(CDec(2.5), 23)
        Case 12: v = Round(CByte(1), 23)
        Case 13: v = Round(True, 23)
        Case 14: v = Round(Empty, 23)
        Case 15: v = Round(Null, 23)
        Case 16: v = Round(False, 2)
        Case 17: v = Round(CByte(1), -1)
        Case 18: v = Round(CCur(2.5), 23)
        Case 19: v = Round(CCur(2.5), -1)
        Case 20: v = Round(CDec(2.5), -1)
        Case 21: v = Round(True, -1)
        Case 22: v = Round(Empty, -1)
        Case 23: v = Round(CSng(2.5), 23)
        Case 24: v = Round(CDate(2.5), 23)
    End Select
    T = "RoundEdges." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "RoundEdges." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
