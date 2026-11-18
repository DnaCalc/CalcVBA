Attribute VB_Name = "P_VarType"
Option Explicit
' VarType (Information; includes vbArray combinations).
Private Const N As Long = 65

Private Function HM(Optional x As Variant) As Long
    HM = VarType(x)
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    Dim by As Byte, i As Integer, l As Long, sg As Single, d As Double, cu As Currency, dt As Date, s As String, bo As Boolean
    Dim va As Variant, o As Object, c As Collection, fa(1 To 3) As Long, da() As Double, ua() As String, m(1 To 2, 1 To 2) As Integer
    Dim ov() As Object, vf(1 To 2) As Variant, bv() As Byte, dic As Object, sa() As Single, ba() As Boolean, ca() As Currency, dta() As Date, ia() As Integer
    On Error GoTo fail
    Select Case k
        Case 1: v = VarType(by)                       ' vbByte 17
        Case 2: v = VarType(i)                        ' vbInteger 2
        Case 3: v = VarType(l)                        ' vbLong 3
        Case 4: v = VarType(sg)                       ' vbSingle 4
        Case 5: v = VarType(d)                        ' vbDouble 5
        Case 6: v = VarType(cu)                       ' vbCurrency 6
        Case 7: v = VarType(dt)                       ' vbDate 7
        Case 8: v = VarType(s)                        ' vbString 8
        Case 9: v = VarType(bo)                       ' vbBoolean 11
        Case 10: v = VarType(va)                       ' uninitialised Variant: vbEmpty 0
        Case 11: va = Null: v = VarType(va)            ' vbNull 1
        Case 12: va = CVErr(5): v = VarType(va)        ' vbError 10
        Case 13: va = CDec(1): v = VarType(va)         ' vbDecimal 14
        Case 14: v = VarType(o)                        ' unset Object variable: vbObject 9
        Case 15: Set c = New Collection: v = VarType(c)   ' vbObject
        Case 16: Set dic = CreateObject("Scripting.Dictionary"): v = VarType(dic)   ' vbObject
        Case 17: Set va = Nothing: v = VarType(va)     ' Variant holding Nothing
        Case 18: v = VarType(Nothing)                  ' Nothing literal
        Case 19: v = VarType(Array(1, 2))              ' vbArray + vbVariant
        Case 20: v = VarType(Array())                  ' empty Variant array
        Case 21: v = VarType(fa)                       ' vbArray + vbLong
        Case 22: v = VarType(da)                       ' unallocated Double array
        Case 23: v = VarType(ua)                       ' unallocated String array
        Case 24: v = VarType(m)                        ' 2-D Integer array
        Case 25: v = VarType(ov)                       ' Object array
        Case 26: v = VarType(vf)                       ' fixed Variant array
        Case 27: v = VarType(bv)                       ' Byte array
        Case 28: v = VarType(sa)                       ' Single array
        Case 29: v = VarType(ba)                       ' Boolean array
        Case 30: v = VarType(ca)                       ' Currency array
        Case 31: v = VarType(dta)                      ' Date array
        Case 32: v = VarType(ia)                       ' Integer array
        Case 33: va = fa: v = VarType(va)              ' Variant holding a Long array
        Case 34: v = VarType(Split("a b"))             ' Split: String array
        Case 35: v = VarType(Array(Array(1)))          ' nested array
        Case 36: v = VarType(Array(1)(0))              ' element
        Case 37: v = VarType(1)                        ' Integer literal
        Case 38: v = VarType(40000)                    ' Long literal
        Case 39: v = VarType(1.5)                      ' Double literal
        Case 40: v = VarType(1@)                       ' Currency literal
        Case 41: v = VarType(#1/1/2000#)               ' Date literal
        Case 42: v = VarType("")                       ' String literal
        Case 43: v = VarType(True)                     ' Boolean literal
        Case 44: v = VarType(Null)                     ' Null literal
        Case 45: v = VarType(Empty)                    ' Empty literal
        Case 46: v = VarType(CVErr(0))                 ' Error literal
        Case 47: v = VarType(HM())                     ' result of the helper: Long
        Case 48: v = HM()                              ' Missing argument: vbError 10
        Case 49: v = VarType(Err)                      ' ErrObject: vbObject
        Case 50: v = VarType(CByte(1) + CByte(1))      ' Byte + Byte
        Case 51: v = VarType(1 + 1.5)                  ' Double
        Case 52: v = VarType(Null + 1)                 ' Null
        Case 53: v = VarType(Empty + Empty)            ' Empty + Empty: Integer
        Case 54: v = VarType(Empty & Empty)            ' String
        Case 55: v = vbEmpty & "," & vbNull & "," & vbInteger & "," & vbLong & "," & vbSingle & "," & vbDouble & "," & vbCurrency & "," & vbDate   ' constants
        Case 56: v = vbString & "," & vbObject & "," & vbError & "," & vbBoolean & "," & vbVariant & "," & vbDataObject & "," & vbDecimal & "," & vbByte   ' constants
        Case 57: v = vbLongLong & "," & vbUserDefinedType & "," & vbArray   ' constants
        Case 58: v = VarType(va) = vbEmpty             ' comparison with the constant
        Case 59: v = (VarType(fa) And vbArray) = vbArray    ' array bit test
        Case 60: v = VarType(fa) And Not vbArray       ' element type of an array
        Case 61: v = TypeName(VarType(1))              ' VarType returns Integer
        Case 62: v = VarType(VarName:=1)               ' named argument
        Case 63: v = VarType(CLng(5) * 2)              ' Long * Integer
        Case 64: v = VarType(Chr(65))                  ' Variant String
        Case 65: v = VarType(Len("a"))                 ' Long
    End Select
    T = "VarType." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "VarType." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
