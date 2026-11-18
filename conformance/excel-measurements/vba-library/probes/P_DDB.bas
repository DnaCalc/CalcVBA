Attribute VB_Name = "P_DDB"
Option Explicit
' DDB (VBA.Financial).
Private Const N As Long = 25

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = DDB(2400, 300, 10, 1)                             ' first period
        Case 2: v = DDB(2400, 300, 10, 2)                             ' second period
        Case 3: v = DDB(2400, 300, 10, 10)                            ' last period
        Case 4: v = DDB(2400, 300, 10, 5)                             ' middle
        Case 5: v = DDB(2400, 300, 10, 1, 1.5)                        ' factor 1.5
        Case 6: v = DDB(2400, 300, 10, 3, 3)                          ' factor 3
        Case 7: v = DDB(2400, 300, 10, 1, 0)                          ' factor 0: error 5
        Case 8: v = DDB(2400, 300, 10, 1, -1)                         ' negative factor: error 5
        Case 9: v = DDB(2400, 300, 10, 11)                            ' period past life: error 5
        Case 10: v = DDB(2400, 300, 10, 0)                            ' period 0: error 5
        Case 11: v = DDB(2400, 300, 0, 1)                             ' life 0: error 5
        Case 12: v = DDB(2400, 3000, 10, 1)                           ' salvage above cost
        Case 13: v = DDB(2400, -300, 10, 1)                           ' negative salvage
        Case 14: v = DDB(-2400, 300, 10, 1)                           ' negative cost
        Case 15: v = DDB(0, 0, 10, 1)                                 ' zero cost
        Case 16: v = DDB(2400, 300, 10, 1.5)                          ' fractional period
        Case 17: v = DDB(2400, 300, 10.5, 2)                          ' fractional life
        Case 18: v = DDB(1000, 0, 5, 5)                               ' zero salvage
        Case 19: v = DDB(1, 0, 3, 2)                                  ' small
        Case 20: v = DDB(2400, 300, 1, 1)                             ' life 1
        Case 21: v = DDB(1000.55, 100.25, 7, 4, 2.5)                  ' awkward decimals
        Case 22: v = DDB(2400, 300, 10, 1, Null)                      ' Null factor
        Case 23: v = "compile-refused"
        Case 24: v = DDB(Cost:=2400, Salvage:=300, Life:=10, Period:=1, Factor:=2)' named
        Case 25: v = DDB(1E+15, 1, 12, 6)                             ' large cost
    End Select
    T = "DDB." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "DDB." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
