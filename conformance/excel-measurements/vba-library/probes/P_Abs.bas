Attribute VB_Name = "P_Abs"
Option Explicit
' Abs (VBA.Math).
Private Const N As Long = 27

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    Dim i As Integer, l As Long, c As Currency, a As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Abs(-3)                                           ' Integer literal
        Case 2: v = Abs(-3.5)                                         ' Double
        Case 3: v = Abs(0)                                            ' zero
        Case 4: v = Abs(-0#)                                          ' negative zero Double
        Case 5: i = -5: v = Abs(i)                                    ' Integer variable
        Case 6: i = -32768: v = Abs(i)                                ' most negative Integer: overflow error 6
        Case 7: l = -70000: v = Abs(l)                                ' Long
        Case 8: l = -2147483647 - 1: v = Abs(l)                       ' most negative Long: overflow error 6
        Case 9: v = Abs(CCur(-1.5))                                   ' Currency
        Case 10: c = -922337203685477.5807@: c = c - 0.0001@: v = Abs(c)' most negative Currency
        Case 11: v = Abs(CDec(-1.25))                                 ' Decimal
        Case 12: v = Abs(CSng(-2.5))                                  ' Single
        Case 13: v = Abs(CByte(5))                                    ' Byte
        Case 14: v = Abs(CDate(-5))                                   ' Date
        Case 15: v = Abs(CBool(True))                                 ' Boolean True
        Case 16: v = Abs(Null)                                        ' Null propagates
        Case 17: v = Abs(Empty)                                       ' Empty
        Case 18: v = Abs("-3.5")                                      ' numeric string gives Double
        Case 19: v = Abs("abc")                                       ' error 13
        Case 20: v = Abs("")                                          ' empty string: error 13
        Case 21: v = Abs(-1E+300)                                     ' huge
        Case 22: v = Abs(CVErr(5))                                    ' Error value: error 13
        Case 23: a = Array(1): v = Abs(a)                             ' array: error 13
        Case 24: v = "compile-refused"
        Case 25: v = Abs(CDec("-79228162514264337593543950335"))      ' Decimal max
        Case 26: v = Abs(-32768)                                      ' Integer literal -32768
        Case 27: v = Abs(-2147483648#)                                ' Double literal at Long edge
    End Select
    T = "Abs." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Abs." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
