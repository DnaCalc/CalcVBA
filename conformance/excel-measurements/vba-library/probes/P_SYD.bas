Attribute VB_Name = "P_SYD"
Option Explicit
' SYD (VBA.Financial).
Private Const N As Long = 17

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = SYD(30000, 7500, 10, 1)                           ' first period
        Case 2: v = SYD(30000, 7500, 10, 10)                          ' last period
        Case 3: v = SYD(30000, 7500, 10, 5)                           ' middle
        Case 4: v = SYD(30000, 7500, 10, 0)                           ' period 0: error 5
        Case 5: v = SYD(30000, 7500, 10, 11)                          ' period past life: error 5
        Case 6: v = SYD(30000, 7500, 0, 1)                            ' life 0: error 5
        Case 7: v = SYD(1000, 100, 5, 2.5)                            ' fractional period
        Case 8: v = SYD(1000, 100, 5.5, 2)                            ' fractional life
        Case 9: v = SYD(100, 0, 3, 1)                                 ' zero salvage
        Case 10: v = SYD(1, 0, 1, 1)                                  ' life 1
        Case 11: v = SYD(0.1, 0.01, 7, 3)                             ' small values
        Case 12: v = SYD(1000, 2000, 5, 2)                            ' salvage above cost
        Case 13: v = SYD(1000, -100, 5, 2)                            ' negative salvage
        Case 14: v = SYD(-1000, 100, 5, 2)                            ' negative cost
        Case 15: v = SYD(1000.55, 100.25, 7, 4)                       ' awkward decimals
        Case 16: v = SYD(1E+15, 1, 12, 6)                             ' large cost
        Case 17: v = SYD(Cost:=1000, Salvage:=100, Life:=5, Period:=2)' named
    End Select
    T = "SYD." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "SYD." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
