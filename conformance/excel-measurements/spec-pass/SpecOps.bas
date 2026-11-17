Attribute VB_Name = "SpecOps"
Option Explicit
' MS-VBAL §5.6.9 operators: open points from the dc-ez3rd code inspection.

Private Function Esc(ByVal s As String) As String
    Dim i As Long, c As Long
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If c < 32 Or c > 126 Then Esc = Esc & "{" & c & "}" Else Esc = Esc & Mid$(s, i, 1)
    Next
End Function

Private Function Sh(ByVal v As Variant) As String
    If IsObject(v) Then Sh = "obj " & TypeName(v): Exit Function
    If IsArray(v) Then Sh = TypeName(v) & "[" & LBound(v) & "," & UBound(v) & "]": Exit Function
    Select Case VarType(v)
        Case vbNull: Sh = "Null"
        Case vbEmpty: Sh = "Empty"
        Case vbDate: Sh = "Date:" & CStr(CDbl(v))
        Case vbString: Sh = "String:[" & Esc(v) & "]"
        Case Else: Sh = TypeName(v) & ":" & CStr(v)
    End Select
End Function

Private Function T(ByVal k As String) As String
    Dim v As Variant, vv As Variant, cu As Currency, dd As Date, s As String, a As Variant, b As Variant
    Dim sg As Single, d0 As Double, bb() As Byte, vn As Variant
    On Error GoTo fail
    Select Case k
        Case "negCur": cu = 1.5: vv = cu: v = -vv
        Case "negDateVar": vv = #1/2/2000#: v = -vv
        Case "negDateDecl": dd = #1/2/2000#: v = -dd
        Case "negSng": vv = CSng(1.5): v = -vv
        Case "negDec": vv = CDec(1.5): v = -vv
        Case "decAdd": v = CDec(1) + 1
        Case "decMul": v = CDec("0.1") * 3
        Case "decDiv3": v = CDec(1) / 3
        Case "decDiv0": v = CDec(1) / CDec(0)
        Case "decOverflow": v = CDec("79228162514264337593543950335") + 1
        Case "decSubPrec": v = CDec("1.0000000000000000000000000001") - 1
        Case "decIntDiv": v = CDec(7) \ 2
        Case "decMod": v = CDec("7.5") Mod 2
        Case "decPow": v = CDec(2) ^ 2
        Case "emptyEmptyAdd": vv = Empty: v = vv + vv
        Case "emptyEmptyMul": vv = Empty: v = vv * vv
        Case "emptyTrue": vv = Empty: v = vv + True
        Case "emptyMinusTrue": vv = Empty: v = vv - True
        Case "not3e9": vv = 3000000000#: v = Not vv
        Case "and3e9": vv = 3000000000#: v = vv And 1
        Case "orStrBig": vv = "3000000000": v = vv Or 0
        Case "xorCurBig": vv = CCur(1000000000000#): v = vv Xor 1
        Case "catArray": v = Array(1) & "x"
        Case "catErr": v = "x" & CVErr(5)
        Case "nullCatNull": vn = Null: v = vn & vn
        Case "nullCatNullStr": vn = Null: s = vn & vn: v = s
        Case "emptyPlusStr5": vv = Empty: v = vv + "5"
        Case "emptyPlusStrA": vv = Empty: v = vv + "a"
        Case "oneDivEmpty": d0 = 1: v = d0 / Empty
        Case "zeroDivEmpty": d0 = 0: v = d0 / Empty
        Case "dateOverVar": vv = #12/31/9999#: vv = vv + 1: v = vv
        Case "dateOverDecl": dd = #12/31/9999#: v = dd + 1
        Case "intAndNull": v = CInt(0) And Null
        Case "byteAndNull": v = CByte(0) And Null
        Case "llAndNull": v = CLngLng(0) And Null
        Case "fiveOrNull": v = 5 Or Null
        Case "fiveImpNull": v = 5 Imp Null
        Case "nullImpFive": v = Null Imp 5
        Case "strEqTrue": s = "True": v = (s = True)
        Case "strEqDate": s = "1/1/2000": v = (s = #1/1/2000#)
        Case "strEqInt": s = "1.4": v = (s = CInt(1))
        Case "errEqNum": v = (CVErr(2042) = 1)
        Case "errLtErr": v = (CVErr(10) < CVErr(9))
        Case "varBoolStr": a = True: b = "True": v = (a = b)
        Case "varDateStr": a = #1/1/2000#: b = "1/1/2000": v = (a = b)
        Case "sngEqDbl": sg = 0.1: v = (sg = 0.1)
        Case "likeOpenBr": v = ("[" Like "[")
        Case "likeRevRange": v = ("b" Like "[z-a]")
        Case "likeEmptyBr": v = ("ab" Like "a[]b")
        Case Else: v = "unknown case"
    End Select
    T = k & "=" & Sh(v): Exit Function
fail:
    T = k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim keys As Variant, i As Long, s As String
    keys = Split("negCur negDateVar negDateDecl negSng negDec decAdd decMul decDiv3 decDiv0 decOverflow decSubPrec decIntDiv decMod decPow emptyEmptyAdd emptyEmptyMul emptyTrue emptyMinusTrue not3e9 and3e9 orStrBig xorCurBig catArray catErr nullCatNull nullCatNullStr emptyPlusStr5 emptyPlusStrA oneDivEmpty zeroDivEmpty dateOverVar dateOverDecl intAndNull byteAndNull llAndNull fiveOrNull fiveImpNull nullImpFive strEqTrue strEqDate strEqInt errEqNum errLtErr varBoolStr varDateStr sngEqDbl likeOpenBr likeRevRange likeEmptyBr", " ")
    For i = 0 To UBound(keys)
        s = s & T(keys(i)) & vbLf
    Next
    Run = s
End Function
