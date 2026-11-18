Attribute VB_Name = "P_Round"
Option Explicit
' Round (VBA.Math).
Private Const N As Long = 51

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Round(2.5)                                        ' banker: 2
        Case 2: v = Round(3.5)                                        ' banker: 4
        Case 3: v = Round(-2.5)                                       ' banker: -2
        Case 4: v = Round(0.5)                                        ' banker: 0
        Case 5: v = Round(1.5)                                        ' 2
        Case 6: v = Round(-0.5)                                       ' negative zero?
        Case 7: v = Round(2.4)                                        ' down
        Case 8: v = Round(2.6)                                        ' up
        Case 9: v = Round(2.5, 0)                                     ' digits 0
        Case 10: v = Round(2.675, 2)                                  ' binary 2.675 below half
        Case 11: v = Round(1.005, 2)                                  ' binary 1.005 below half
        Case 12: v = Round(0.285, 2)                                  ' 0.285 binary
        Case 13: v = Round(1.45, 1)                                   ' 1.45
        Case 14: v = Round(1.55, 1)                                   ' 1.55
        Case 15: v = Round(0.125, 2)                                  ' exact half: banker 0.12
        Case 16: v = Round(0.375, 2)                                  ' exact half: banker 0.38
        Case 17: v = Round(-0.125, 2)                                 ' negative exact half
        Case 18: v = Round(1234.5678, 2)                              ' typical
        Case 19: v = Round(1234.5678, 3)                              ' typical 3
        Case 20: v = Round(1234.5678, 10)                             ' digits past precision
        Case 21: v = Round(3.14159265358979, 15)                      ' digits 15
        Case 22: v = Round(3.14159265358979, 22)                      ' digits 22 maximum
        Case 23: v = Round(3.14159265358979, 23)                      ' digits 23: past documented range
        Case 24: v = Round(1234.5678, -1)                             ' negative digits: error 5
        Case 25: v = Round(12.5, -1)                                  ' negative digits: error 5
        Case 26: v = Round(CInt(5))                                   ' Integer type
        Case 27: v = Round(CLng(5), 2)                                ' Long type
        Case 28: v = Round(CByte(5))                                  ' Byte type
        Case 29: v = Round(CSng(2.5))                                 ' Single type
        Case 30: v = Round(CSng(0.285), 2)                            ' Single with digits
        Case 31: v = Round(CCur(2.5))                                 ' Currency type
        Case 32: v = Round(CCur(1234.5678), 2)                        ' Currency digits
        Case 33: v = Round(CDec(2.5))                                 ' Decimal type
        Case 34: v = Round(CDec("0.285"), 2)                          ' Decimal digits
        Case 35: v = Round(CDate(2.5))                                ' Date type
        Case 36: v = Round(True)                                      ' Boolean
        Case 37: v = Round("2.5")                                     ' numeric string
        Case 38: v = Round("abc")                                     ' error 13
        Case 39: v = Round(Null)                                      ' Null
        Case 40: v = Round(Empty)                                     ' Empty
        Case 41: v = Round(2.5, Null)                                 ' Null digits
        Case 42: v = Round(1E+300)                                    ' huge
        Case 43: v = Round(1E+300, 2)                                 ' huge with digits
        Case 44: v = Round(4503599627370497#)                         ' 2^52+1
        Case 45: v = Round(0.49999999999999994)                       ' just below half
        Case 46: v = Round(2147483647.5)                              ' past Long
        Case 47: v = Round(-1E-300, 5)                                ' tiny negative
        Case 48: v = Round(Number:=2.5, NumDigitsAfterDecimal:=0)     ' named arguments
        Case 49: v = Round(2.5, "1")                                  ' numeric string digits
        Case 50: v = Round(2.5, 1.5)                                  ' fractional digits
        Case 51: v = Round(2.5, 2.5)                                  ' fractional digits 2.5
    End Select
    T = "Round." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Round." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
