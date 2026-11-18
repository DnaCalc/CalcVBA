Attribute VB_Name = "P_SLN"
Option Explicit
' SLN (VBA.Financial).
Private Const N As Long = 16

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = SLN(30000, 7500, 10)                              ' typical
        Case 2: v = SLN(100, 0, 3)                                    ' non-terminating
        Case 3: v = SLN(10, 1, 3)                                     ' 10/1/3
        Case 4: v = SLN(0.1, 0.01, 3)                                 ' small
        Case 5: v = SLN(30000, 30000, 10)                             ' cost equals salvage
        Case 6: v = SLN(30000, 40000, 10)                             ' salvage above cost
        Case 7: v = SLN(30000, -7500, 10)                             ' negative salvage
        Case 8: v = SLN(-30000, 7500, 10)                             ' negative cost
        Case 9: v = SLN(30000, 7500, 0)                               ' life 0
        Case 10: v = SLN(30000, 7500, -5)                             ' negative life
        Case 11: v = SLN(30000, 7500, 2.5)                            ' fractional life
        Case 12: v = SLN(1E+300, 1, 7)                                ' huge
        Case 13: v = SLN(1.7976931348623157E+308, -1.7976931348623157E+308, 1)' overflow
        Case 14: v = SLN(1000.55, 100.25, 7)                          ' awkward decimals
        Case 15: v = SLN(Cost:=100, Salvage:=10, Life:=3)             ' named
        Case 16: v = SLN("100", "10", "3")                            ' numeric strings
    End Select
    T = "SLN." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "SLN." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
