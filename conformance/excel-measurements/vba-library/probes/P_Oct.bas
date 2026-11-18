Attribute VB_Name = "P_Oct"
Option Explicit
' Oct / Oct$ (64-bit VBA: uses LongLong) (VBA.Conversion; MS-VBAL 6.1.2.x).
Private Const N As Long = 58

Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, c As Currency, dt As Date, ds As Double, sg As Single
    Dim l As Long, i As Integer, by As Byte, b As Boolean, s As String, o As Object
    On Error GoTo fail
#If Win64 Then
    Select Case k
        Case 1: v = Oct(0)
        Case 2: v = Oct(255)
        Case 3: v = Oct(256)
        Case 4: v = Oct(15.5)   ' banker rounding to 16
        Case 5: v = Oct(14.5)   ' banker rounding to 14
        Case 6: v = Oct(255.5)   ' rounds to 256
        Case 7: v = Oct(-0.5)   ' rounds to 0
        Case 8: v = Oct(-1)   ' Integer width
        Case 9: v = Oct(-1&)   ' Long width
        Case 10: v = Oct(-1^)   ' LongLong width
        Case 11: v = Oct(-32768)   ' Integer min
        Case 12: v = Oct(-32769)   ' Long
        Case 13: v = Oct(32767)
        Case 14: v = Oct(32768)   ' Long
        Case 15: v = Oct(65535)   ' Long
        Case 16: v = Oct(65536)
        Case 17: v = Oct(2147483647)
        Case 18: v = Oct(-2147483648#)   ' Double within Long
        Case 19: v = Oct(2147483648#)   ' Double beyond Long
        Case 20: v = Oct(4294967295#)   ' Double beyond Long
        Case 21: v = Oct(1E10)   ' Double
        Case 22: v = Oct(1E20)   ' beyond LongLong
        Case 23: v = Oct(9223372036854775807^)   ' LongLong max
        Case 24: v = Oct(-9223372036854775807^ - 1)   ' LongLong min
        Case 25: v = Oct(CByte(255))   ' Byte
        Case 26: v = Oct(CByte(0))
        Case 27: v = Oct(CInt(-2))   ' Integer
        Case 28: v = Oct(CLng(-2))   ' Long
        Case 29: v = Oct(CSng(255.5))   ' Single
        Case 30: v = Oct(CCur(255.5))   ' Currency
        Case 31: v = Oct(CDec(255))   ' Decimal
        Case 32: v = Oct(CDec("1E20"))   ' Decimal beyond range
        Case 33: v = Oct(True)   ' Boolean -1
        Case 34: v = Oct(False)
        Case 35: v = Oct("255")   ' numeric string
        Case 36: v = Oct("&HFF")   ' hex string
        Case 37: v = Oct("&O17")   ' octal string
        Case 38: v = Oct("12.5")
        Case 39: v = Oct("abc")   ' 13
        Case 40: v = Oct("")   ' 13
        Case 41: v = Oct(Null)   ' Null
        Case 42: v = Oct(Empty)   ' zero
        Case 43: v = Oct(CVErr(5))   ' 13
        Case 44: v = Oct(#1/2/2020#)   ' Date
        Case 45: v = Oct(Number:=255)   ' named argument
        Case 46: v = Oct(Array(1))   ' 13
        Case 47: v = Oct(-2147483649#)   ' below Long min
        Case 48: v = Oct(&H7FFFFFFF&)
        Case 49: v = Oct(&H80000000)
        Case 50: v = Oct(&HFFFF)   ' Integer literal -1
        Case 51: v = Oct(&HFFFF&)   ' Long literal 65535
        Case 52: v = TypeName(Oct(5))   ' Variant String
        Case 53: v = Oct$(255)   ' String
        Case 54: v = Oct$(-1)   ' $ form negative Integer
        Case 55: v = Oct$(255.5)   ' $ form Double
        Case 56: v = Oct$(Null)   ' 94
        Case 57: v = Oct$("12")   ' $ form string
        Case 58: v = Oct$(Number:=255)   ' $ named
    End Select
#Else
    v = "n/a-32bit"
#End If
    T = "Oct." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Oct." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
