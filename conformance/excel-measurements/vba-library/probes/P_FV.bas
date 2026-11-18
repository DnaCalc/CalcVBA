Attribute VB_Name = "P_FV"
Option Explicit
' FV (VBA.Financial).
Private Const N As Long = 18

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = FV(0.05 / 12, 360, -500)                          ' savings
        Case 2: v = FV(0.05 / 12, 360, -500, 0, 1)                    ' due
        Case 3: v = FV(0, 12, -100, -50)                              ' zero rate
        Case 4: v = FV(0, 12, -100, -50, 1)                           ' zero rate due
        Case 5: v = FV(0.1, 10, -100, -1000, 1)                       ' pv and due
        Case 6: v = FV(0.1, 10, 100, 1000)                            ' positive pmt
        Case 7: v = FV(0.1, 0, -100)                                  ' nper 0
        Case 8: v = FV(0.1, 2.5, -100)                                ' fractional nper
        Case 9: v = FV(-0.01, 12, -100)                               ' negative rate
        Case 10: v = FV(0.07, 40, -1000.55)                           ' awkward decimals
        Case 11: v = FV(1E-10, 360, -1)                               ' tiny rate
        Case 12: v = FV(5, 50, -1)                                    ' large growth
        Case 13: v = FV(0.5, 1000, -1)                                ' overflow
        Case 14: v = FV(-1, 5, -100, -50)                             ' rate -1
        Case 15: v = FV(0.05, -10, -100)                              ' negative nper
        Case 16: v = FV(0.1, 10, -100, Null)                          ' Null pv
        Case 17: v = FV(0.1, 10, -100, , 1)                           ' omitted pv
        Case 18: v = FV(Rate:=0.1, NPer:=10, Pmt:=-100, PV:=0, Due:=0)' named
    End Select
    T = "FV." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "FV." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
