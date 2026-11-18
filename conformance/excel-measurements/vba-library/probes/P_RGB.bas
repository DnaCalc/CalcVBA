Attribute VB_Name = "P_RGB"
Option Explicit
' RGB (Information; Red + Green*256 + Blue*65536, components above 255 clamp to 255).
Private Const N As Long = 31

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    Dim i As Integer, b As Byte
    On Error GoTo fail
    Select Case k
        Case 1: v = RGB(0, 0, 0)                      ' black
        Case 2: v = RGB(255, 0, 0)                    ' red
        Case 3: v = RGB(0, 255, 0)                    ' green
        Case 4: v = RGB(0, 0, 255)                    ' blue
        Case 5: v = RGB(255, 255, 255)                ' white
        Case 6: v = RGB(1, 2, 3)                      ' component order
        Case 7: v = RGB(128, 64, 32)
        Case 8: v = RGB(256, 0, 0)                    ' clamp above 255
        Case 9: v = RGB(0, 300, 0)
        Case 10: v = RGB(0, 0, 1000)
        Case 11: v = RGB(32767, 32767, 32767)          ' largest Integer
        Case 12: v = RGB(32768, 0, 0)                  ' overflow: error 6
        Case 13: v = RGB(-1, 0, 0)                     ' negative: error 5
        Case 14: v = RGB(0, -1, 0)
        Case 15: v = RGB(0, 0, -1)
        Case 16: v = RGB(-32768, 0, 0)
        Case 17: v = RGB(1.5, 2.5, 3.5)                ' rounding of fractions
        Case 18: v = RGB(254.5, 0, 0)                  ' rounds to 254
        Case 19: v = RGB(255.5, 0, 0)                  ' rounds to 256 then clamps
        Case 20: v = RGB(-0.4, 0, 0)                   ' rounds to 0
        Case 21: v = RGB("255", "0", "0")              ' numeric strings
        Case 22: v = RGB("x", 0, 0)                    ' error 13
        Case 23: v = RGB(Null, 0, 0)                   ' error 94
        Case 24: v = RGB(0, Null, 0)
        Case 25: v = RGB(Empty, Empty, Empty)          ' zeros
        Case 26: v = RGB(True, 0, 0)                   ' -1: error 5
        Case 27: v = RGB(False, False, False)
        Case 28: b = 200: i = 100: v = RGB(b, i, b)    ' Byte and Integer variables
        Case 29: v = TypeName(RGB(1, 2, 3))            ' result type: Long
        Case 30: v = RGB(Red:=1, Green:=2, Blue:=3)    ' named arguments
        Case 31: v = RGB(Blue:=255, Red:=0, Green:=0)  ' named in a different order
    End Select
    T = "RGB." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "RGB." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
