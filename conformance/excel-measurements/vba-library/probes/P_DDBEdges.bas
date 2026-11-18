Attribute VB_Name = "P_DDBEdges"
Option Explicit
' DDB short lives, salvage and factor boundaries measured in Excel.
Private Const N As Long = 16

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = DDB(100, 10, 2, 1, 1)
        Case 2: v = DDB(100, 10, 2, 2, 1)
        Case 3: v = DDB(100, 10, 2, 0.5, 3)
        Case 4: v = DDB(100, 10, 2, 1.5, 3)
        Case 5: v = DDB(100, 10, 1.5, 1.25, 1)
        Case 6: v = DDB(100, 110, 1, 1, 1)
        Case 7: v = DDB(100, 110, 3, 0.5, 1)
        Case 8: v = DDB(100, 10, 0.5, 0.25, 1)
        Case 9: v = DDB(0, 10, 3, 1, 1)
        Case 10: v = DDB(-100, 10, 3, 1, 1)
        Case 11: v = DDB(-100, -10, 3, 1, 1)
        Case 12: v = DDB(-100, 10, 3, 1, 0)
        Case 13: v = DDB(100, 10, 3, 2, 6)
        Case 14: v = DDB(100, 10, 3, 1.5, 6)
        Case 15: v = DDB(100, 10, 3, 3, 6)
        Case 16: v = DDB(100, 10, 3, 2, 3)
    End Select
    T = "DDBEdges." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "DDBEdges." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
