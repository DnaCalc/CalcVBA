Attribute VB_Name = "P_Rate"
Option Explicit
' Rate (VBA.Financial).
Private Const N As Long = 8

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Rate(360, -1073.64, 200000)                       ' 30-year loan
        Case 2: v = Rate(360, -1073.64, 200000, 0, 1)                 ' due
        Case 3: v = Rate(60, -200, 10000, 500)                        ' fv
        Case 4: v = Rate(12, -100, 1000, 0, 0, 0.5)                   ' guess 0.5
        Case 5: v = Rate(10, -100, 1000)                              ' zero rate exactly
        Case 6: v = Rate(0, -100, 1000)                               ' nper 0
        Case 7: v = Rate(10, 100, 1000)                               ' non-convergent: error 5
        Case 8: v = Rate(1, -110, 100)                                ' single period
    End Select
    T = "Rate." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Rate." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
