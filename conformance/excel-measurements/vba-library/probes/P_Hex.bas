Attribute VB_Name = "P_Hex"
Option Explicit
' Hex / Hex$ (64-bit VBA: uses LongLong) (VBA.Conversion; MS-VBAL 6.1.2.x).
Private Const N As Long = 58

Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, c As Currency, dt As Date, ds As Double, sg As Single
    Dim l As Long, i As Integer, by As Byte, b As Boolean, s As String, o As Object
    On Error GoTo fail
#If Win64 Then
    Select Case k
        Case 1: v = Hex(0)
        Case 2: v = Hex(255)
        Case 3: v = Hex(256)
        Case 4: v = Hex(15.5)   ' banker rounding to 16
        Case 5: v = Hex(14.5)   ' banker rounding to 14
        Case 6: v = Hex(255.5)   ' rounds to 256
        Case 7: v = Hex(-0.5)   ' rounds to 0
        Case 8: v = Hex(-1)   ' Integer width
        Case 9: v = Hex(-1&)   ' Long width
        Case 10: v = Hex(-1^)   ' LongLong width
        Case 11: v = Hex(-32768)   ' Integer min
        Case 12: v = Hex(-32769)   ' Long
        Case 13: v = Hex(32767)
        Case 14: v = Hex(32768)   ' Long
        Case 15: v = Hex(65535)   ' Long
        Case 16: v = Hex(65536)
        Case 17: v = Hex(2147483647)
        Case 18: v = Hex(-2147483648#)   ' Double within Long
        Case 19: v = Hex(2147483648#)   ' Double beyond Long
        Case 20: v = Hex(4294967295#)   ' Double beyond Long
        Case 21: v = Hex(1E10)   ' Double
        Case 22: v = Hex(1E20)   ' beyond LongLong
        Case 23: v = Hex(9223372036854775807^)   ' LongLong max
        Case 24: v = Hex(-9223372036854775807^ - 1)   ' LongLong min
        Case 25: v = Hex(CByte(255))   ' Byte
        Case 26: v = Hex(CByte(0))
        Case 27: v = Hex(CInt(-2))   ' Integer
        Case 28: v = Hex(CLng(-2))   ' Long
        Case 29: v = Hex(CSng(255.5))   ' Single
        Case 30: v = Hex(CCur(255.5))   ' Currency
        Case 31: v = Hex(CDec(255))   ' Decimal
        Case 32: v = Hex(CDec("1E20"))   ' Decimal beyond range
        Case 33: v = Hex(True)   ' Boolean -1
        Case 34: v = Hex(False)
        Case 35: v = Hex("255")   ' numeric string
        Case 36: v = Hex("&HFF")   ' hex string
        Case 37: v = Hex("&O17")   ' octal string
        Case 38: v = Hex("12.5")
        Case 39: v = Hex("abc")   ' 13
        Case 40: v = Hex("")   ' 13
        Case 41: v = Hex(Null)   ' Null
        Case 42: v = Hex(Empty)   ' zero
        Case 43: v = Hex(CVErr(5))   ' 13
        Case 44: v = Hex(#1/2/2020#)   ' Date
        Case 45: v = Hex(Number:=255)   ' named argument
        Case 46: v = Hex(Array(1))   ' 13
        Case 47: v = Hex(-2147483649#)   ' below Long min
        Case 48: v = Hex(&H7FFFFFFF&)
        Case 49: v = Hex(&H80000000)
        Case 50: v = Hex(&HFFFF)   ' Integer literal -1
        Case 51: v = Hex(&HFFFF&)   ' Long literal 65535
        Case 52: v = TypeName(Hex(5))   ' Variant String
        Case 53: v = Hex$(255)   ' String
        Case 54: v = Hex$(-1)   ' $ form negative Integer
        Case 55: v = Hex$(255.5)   ' $ form Double
        Case 56: v = Hex$(Null)   ' 94
        Case 57: v = Hex$("12")   ' $ form string
        Case 58: v = Hex$(Number:=255)   ' $ named
    End Select
#Else
    v = "n/a-32bit"
#End If
    T = "Hex." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Hex." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
