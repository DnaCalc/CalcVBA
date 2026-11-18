Attribute VB_Name = "P_Sgn"
Option Explicit
' Sgn (VBA.Math).
Private Const N As Long = 21

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    Dim a As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Sgn(5)                                            ' positive Integer
        Case 2: v = Sgn(-5)                                           ' negative Integer
        Case 3: v = Sgn(0)                                            ' zero
        Case 4: v = Sgn(-0#)                                          ' negative zero Double
        Case 5: v = Sgn(-0.0001)                                      ' small negative
        Case 6: v = Sgn(1E-300)                                       ' tiny positive
        Case 7: v = Sgn(-1E+300)                                      ' huge negative
        Case 8: v = Sgn(CCur(-2))                                     ' Currency
        Case 9: v = Sgn(CDec(3))                                      ' Decimal
        Case 10: v = Sgn(CSng(-0.5))                                  ' Single
        Case 11: v = Sgn(CLng(70000))                                 ' Long
        Case 12: v = Sgn(CByte(7))                                    ' Byte
        Case 13: v = Sgn(CDate(-1))                                   ' Date
        Case 14: v = Sgn(True)                                        ' Boolean True
        Case 15: v = Sgn(Null)                                        ' Null: error 94
        Case 16: v = Sgn(Empty)                                       ' Empty
        Case 17: v = Sgn("-3")                                        ' numeric string
        Case 18: v = Sgn("abc")                                       ' error 13
        Case 19: v = Sgn("")                                          ' empty string: error 13
        Case 20: a = Array(1): v = Sgn(a)                             ' array: error 13
        Case 21: v = "compile-refused"
    End Select
    T = "Sgn." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Sgn." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
