Attribute VB_Name = "P_PV"
Option Explicit
' PV (VBA.Financial).
Private Const N As Long = 14

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = PV(0.05 / 12, 360, -1073.64)                      ' loan principal
        Case 2: v = PV(0.05 / 12, 360, -1073.64, 0, 1)                ' due
        Case 3: v = PV(0, 12, -100)                                   ' zero rate
        Case 4: v = PV(0, 12, -100, -50, 1)                           ' zero rate fv due
        Case 5: v = PV(0.08, 5, -1000, 5000)                          ' fv
        Case 6: v = PV(0.1, 0, -100)                                  ' nper 0
        Case 7: v = PV(0.1, 2.5, -100)                                ' fractional nper
        Case 8: v = PV(-0.02, 24, -100)                               ' negative rate
        Case 9: v = PV(1E-10, 360, -1)                                ' tiny rate
        Case 10: v = PV(0.0725 / 12, 48, -1234.56, 100.5)             ' awkward decimals
        Case 11: v = PV(-1, 5, -100)                                  ' rate -1
        Case 12: v = PV(0.1, 10, 0, 1000)                             ' only fv
        Case 13: v = PV(0.1, 10, -100, Null)                          ' Null fv
        Case 14: v = PV(Rate:=0.1, NPer:=10, Pmt:=-100, FV:=0, Due:=1)' named
    End Select
    T = "PV." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "PV." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
