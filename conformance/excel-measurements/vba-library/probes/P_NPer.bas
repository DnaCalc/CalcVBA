Attribute VB_Name = "P_NPer"
Option Explicit
' NPer (VBA.Financial).
Private Const N As Long = 17

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = NPer(0.05 / 12, -200, 10000)                      ' typical
        Case 2: v = NPer(0.05 / 12, -200, 10000, 0, 1)                ' due
        Case 3: v = NPer(0, -100, 1000)                               ' zero rate
        Case 4: v = NPer(0, -100, 1000, 500)                          ' zero rate fv
        Case 5: v = NPer(0.1, -100, 10000)                            ' payment below interest: error 5
        Case 6: v = NPer(0.1, 0, 1000)                                ' pmt 0
        Case 7: v = NPer(0.1, -100, 1000, 500)                        ' fv
        Case 8: v = NPer(0.1, -100, 1000, 0, 1)                       ' due
        Case 9: v = NPer(0, 0, 1000)                                  ' zero rate zero pmt
        Case 10: v = NPer(-0.1, -100, 1000)                           ' negative rate
        Case 11: v = NPer(0.01, -10, 1000)                            ' pmt below interest
        Case 12: v = NPer(0.075 / 12, -1234.56, 25000.55, 100.25)     ' awkward decimals
        Case 13: v = NPer(1E-12, -100, 1000)                          ' tiny rate
        Case 14: v = NPer(0.1, -100, -1000)                           ' same signs
        Case 15: v = NPer(-1, -100, 1000)                             ' rate -1
        Case 16: v = NPer(0.1, -100, 1000, Null)                      ' Null fv
        Case 17: v = NPer(Rate:=0.05, Pmt:=-100, PV:=1000, FV:=0, Due:=0)' named
    End Select
    T = "NPer." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "NPer." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
