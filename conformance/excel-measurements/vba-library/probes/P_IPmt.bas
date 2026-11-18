Attribute VB_Name = "P_IPmt"
Option Explicit
' IPmt (VBA.Financial).
Private Const N As Long = 18

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = IPmt(0.05 / 12, 1, 360, 200000)                   ' first period
        Case 2: v = IPmt(0.05 / 12, 2, 360, 200000)                   ' second period
        Case 3: v = IPmt(0.05 / 12, 360, 360, 200000)                 ' last period
        Case 4: v = IPmt(0.05 / 12, 100, 360, 200000, 0, 1)           ' due
        Case 5: v = IPmt(0.05 / 12, 1, 360, 200000, 0, 1)             ' due first period
        Case 6: v = IPmt(0.05 / 12, 2, 360, 200000, 0, 1)             ' due second period
        Case 7: v = IPmt(0, 1, 12, 1200)                              ' zero rate
        Case 8: v = IPmt(0.1, 3, 5, -1000, 200)                       ' fv
        Case 9: v = IPmt(0.1, 1, 1, 100)                              ' single period
        Case 10: v = IPmt(0.1, 0, 5, 1000)                            ' per 0: error 5
        Case 11: v = IPmt(0.1, 6, 5, 1000)                            ' per past nper: error 5
        Case 12: v = IPmt(0.1, -1, 5, 1000)                           ' negative per
        Case 13: v = IPmt(0.1, 1, 0, 1000)                            ' nper 0
        Case 14: v = IPmt(0.1, 1.5, 5, 1000)                          ' fractional per
        Case 15: v = IPmt(-0.05, 2, 10, 1000)                         ' negative rate
        Case 16: v = IPmt(0.0725 / 12, 7, 48, 25000.55, 100.25)       ' awkward decimals
        Case 17: v = IPmt(Rate:=0.1, Per:=1, NPer:=5, PV:=1000, FV:=0, Due:=0)' named
        Case 18: v = IPmt(0.1, 1, 5, 1000, Null)                      ' Null fv
    End Select
    T = "IPmt." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "IPmt." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
