Attribute VB_Name = "P_Pmt"
Option Explicit
' Pmt (VBA.Financial).
Private Const N As Long = 21

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Pmt(0.05 / 12, 360, 200000)                       ' 30-year loan
        Case 2: v = Pmt(0.05 / 12, 360, -200000)                      ' negative present value
        Case 3: v = Pmt(0.05 / 12, 360, 200000, 0, 1)                 ' annuity due
        Case 4: v = Pmt(0.08 / 12, 60, -10000, 20000)                 ' future value
        Case 5: v = Pmt(0.08 / 12, 60, -10000, 20000, 1)              ' future value, due
        Case 6: v = Pmt(0, 12, 1200)                                  ' zero rate
        Case 7: v = Pmt(0, 10, 1000, 500, 1)                          ' zero rate, fv, due
        Case 8: v = Pmt(0.1, 1, 100)                                  ' one period
        Case 9: v = Pmt(0.075 / 12, 48, 25000.55, 1000.25)            ' awkward decimals
        Case 10: v = Pmt(1E-12, 360, 1000)                            ' tiny rate
        Case 11: v = Pmt(0.1, 0, 100)                                 ' nper 0
        Case 12: v = Pmt(-0.5, 10, 1000)                              ' negative rate
        Case 13: v = Pmt(0.05, -10, 100)                              ' negative nper
        Case 14: v = Pmt(0.05, 2.5, 100)                              ' fractional nper
        Case 15: v = Pmt(100, 5, 1000)                                ' huge rate
        Case 16: v = Pmt(0.05, 10, 100, Null)                         ' Null fv
        Case 17: v = Pmt(0.05, 10, 100, , 1)                          ' omitted fv
        Case 18: v = Pmt(0.05, 10, 100, 0, 2)                         ' due other than 0/1
        Case 19: v = Pmt(0.05, 10, 100, 0, -1)                        ' due -1
        Case 20: v = Pmt(Rate:=0.01, NPer:=12, PV:=1000, FV:=0, Due:=0)' named arguments
        Case 21: v = Pmt("0.01", "12", "1000")                        ' numeric strings
    End Select
    T = "Pmt." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Pmt." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
