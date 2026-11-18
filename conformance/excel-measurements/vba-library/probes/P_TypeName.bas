Attribute VB_Name = "P_TypeName"
Option Explicit
' TypeName (Information).
Private Const N As Long = 85

Private Function HM(Optional x As Variant) As String
    HM = TypeName(x)
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    Dim by As Byte, i As Integer, l As Long, sg As Single, d As Double, cu As Currency, dt As Date, s As String, bo As Boolean
    Dim va As Variant, o As Object, c As Collection, fa(1 To 3) As Long, da() As Double, ua() As String, m(1 To 2, 1 To 2) As Integer
    Dim ov() As Object, vf(1 To 2) As Variant, fs As String * 5, bv() As Byte, dic As Object
    On Error GoTo fail
    Select Case k
        Case 1: v = TypeName(by)                      ' Byte
        Case 2: v = TypeName(i)                       ' Integer
        Case 3: v = TypeName(l)                       ' Long
        Case 4: v = TypeName(sg)                      ' Single
        Case 5: v = TypeName(d)                       ' Double
        Case 6: v = TypeName(cu)                      ' Currency
        Case 7: v = TypeName(dt)                      ' Date
        Case 8: v = TypeName(s)                       ' String
        Case 9: v = TypeName(bo)                      ' Boolean
        Case 10: v = TypeName(va)                      ' uninitialised Variant: Empty
        Case 11: va = Null: v = TypeName(va)           ' Null
        Case 12: va = Empty: v = TypeName(va)          ' Empty
        Case 13: v = TypeName(o)                       ' unset Object: Nothing
        Case 14: Set c = New Collection: v = TypeName(c)    ' Collection
        Case 15: Set c = New Collection: Set o = c: v = TypeName(o)   ' Object variable holding a Collection
        Case 16: va = CVErr(5): v = TypeName(va)       ' Error
        Case 17: va = CDec(1): v = TypeName(va)        ' Decimal
        Case 18: v = TypeName(Array(1, 2))             ' Variant()
        Case 19: v = TypeName(fa)                      ' Long()
        Case 20: v = TypeName(da)                      ' unallocated Double()
        Case 21: ReDim da(2): v = TypeName(da)         ' allocated Double()
        Case 22: v = TypeName(ua)                      ' String()
        Case 23: v = TypeName(m)                       ' 2-D Integer()
        Case 24: v = TypeName(ov)                      ' Object()
        Case 25: va = fa: v = TypeName(va)             ' Variant holding a Long array
        Case 26: v = TypeName(vf)                      ' fixed Variant()
        Case 27: v = TypeName(fs)                      ' fixed-length String
        Case 28: v = TypeName(bv)                      ' Byte()
        Case 29: v = TypeName(40000)                   ' literal beyond Integer: Long
        Case 30: v = TypeName(1)                       ' Integer literal
        Case 31: v = TypeName(1.5)                     ' Double literal
        Case 32: v = TypeName(1!)                      ' Single literal
        Case 33: v = TypeName(1@)                      ' Currency literal
        Case 34: v = TypeName(1#)                      ' Double literal
        Case 35: v = TypeName(#1/1/2000#)              ' Date literal
        Case 36: v = TypeName("x")                     ' String literal
        Case 37: v = TypeName(True)                    ' Boolean literal
        Case 38: v = TypeName(&HFFFF)                  ' hex literal: Integer (-1)
        Case 39: v = TypeName(&HFFFF&)                 ' hex Long literal
        Case 40: v = TypeName(&H10000)                 ' hex literal needing Long
        Case 41: v = TypeName(1 + 1.5)                 ' mixed arithmetic: Double
        Case 42: v = TypeName(5 \ 2)                   ' integer division: Integer
        Case 43: v = TypeName("1" + 1)                 ' string plus number
        Case 44: v = TypeName(1 / 2)                   ' division: Double
        Case 45: v = TypeName(2 ^ 2)                   ' power: Double
        Case 46: v = TypeName(CInt(1) + CInt(1))       ' Integer + Integer
        Case 47: v = TypeName(CLng(1) + CInt(1))       ' Long + Integer
        Case 48: v = TypeName(CSng(1) + CDbl(1))       ' Single + Double
        Case 49: v = TypeName(CByte(200) + CByte(100)) ' Byte overflow: error 6
        Case 50: v = TypeName(True + 1)                ' Boolean + Integer
        Case 51: v = TypeName(Empty + 1)               ' Empty + Integer
        Case 52: v = TypeName(Empty & "")              ' Empty concatenated
        Case 53: v = TypeName(Null + 1)                ' Null propagates
        Case 54: v = TypeName(Null & "a")              ' Null concatenated
        Case 55: v = TypeName(Nothing)                 ' Nothing literal
        Case 56: v = TypeName(Null)                    ' Null literal
        Case 57: v = TypeName(Empty)                   ' Empty literal
        Case 58: v = TypeName(CVErr(0))                ' Error
        Case 59: v = TypeName(Err)                     ' ErrObject
        Case 60: v = TypeName(TypeName(1))             ' result type
        Case 61: v = TypeName(Chr(65))                 ' Chr returns Variant String
        Case 62: v = TypeName(Chr$(65))                ' Chr$ returns String
        Case 63: v = TypeName(vbNullString)            ' String
        Case 64: v = TypeName(Split(""))               ' Split result: String()
        Case 65: v = TypeName(Array(Array()))          ' nested array
        Case 66: v = TypeName(CVar(1))                 ' CVar Integer
        Case 67: v = TypeName(Now)                     ' Date
        Case 68: v = TypeName(Format(1))               ' Format returns String
        Case 69: v = HM()                              ' Missing is an Error value
        Case 70: v = TypeName(VarName:=1)              ' named argument
        Case 71: Set dic = CreateObject("Scripting.Dictionary"): v = TypeName(dic)   ' Dictionary
        Case 72: v = TypeName(vbYes)                   ' enum constant: Long
        Case 73: v = TypeName(Len("a"))                ' Len returns Long
        Case 74: v = TypeName(UBound(fa))              ' UBound returns Long
        Case 75: v = TypeName(2147483648#)             ' Double literal beyond Long
        Case 76: v = TypeName(Int(2.5))                ' Int of a Double: Double
        Case 77: v = TypeName(Fix(2))                  ' Fix of an Integer: Integer
        Case 78: v = TypeName(CBool(1))                ' Boolean
        Case 79: v = TypeName(CDate(1))                ' Date
        Case 80: v = TypeName(CCur(1))                 ' Currency
        Case 81: v = TypeName(CStr(1))                 ' String
        Case 82: v = TypeName(CLng(1))                 ' Long
        Case 83: v = TypeName(Asc("a"))                ' Integer
        Case 84: v = TypeName(Mid$("abc", 1, 1))       ' String
        Case 85: v = TypeName(Mid("abc", 1, 1))        ' Variant String
    End Select
    T = "TypeName." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "TypeName." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
