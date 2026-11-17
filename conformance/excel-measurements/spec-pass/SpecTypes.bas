Attribute VB_Name = "SpecTypes"
Option Explicit
' MS-VBAL §2 and §5.5 (Let-coercion): open points from the dc-ez3rd code inspection.
Private Type BytT
    b(7) As Byte
End Type
Private Type DblT
    d As Double
End Type

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

Private Function Special(ByVal hi6 As Byte, ByVal hi7 As Byte) As String
    Dim b8 As BytT, d8 As DblT
    b8.b(6) = hi6: b8.b(7) = hi7
    LSet d8 = b8
    Special = CStr(d8.d)
End Function

Private Function T(ByVal k As String) As String
    Dim v As Variant, vv As Variant, by As Byte, dt As Date, s As String, bb() As Byte
    Dim x As Double, ll As LongLong, f As Boolean, c As Currency, o As Object, w As Variant
    On Error GoTo fail
    Select Case k
        Case "byTrue": by = True: v = by
        Case "cbyTrue": v = CByte(True)
        Case "dtBig": dt = 3000000: v = dt
        Case "dtSmall": dt = -700000: v = dt
        Case "cdtBig": v = CDate(3000000)
        Case "strErr": s = CVErr(5): v = s
        Case "bytesToStr": bb = "AB": s = bb: v = s
        Case "varStrToBytes": vv = "AB": bb = vv: v = UBound(bb)
        Case "dbl1d3": x = "1d3": v = x
        Case "dblSpSign": x = "- 5": v = x
        Case "dblThou": x = "1,000": v = x
        Case "dblDollar": x = "$5": v = x
        Case "dblParen": x = "(5)": v = x
        Case "dblTrailMinus": x = "5-": v = x
        Case "dblE400": x = "1E400": v = x
        Case "dblHex": x = "&H10": v = x
        Case "dblPad": x = "  12  ": v = x
        Case "llBigStr": ll = "9007199254740993": v = CStr(ll)
        Case "boolHashTrue": f = "#TRUE#": v = f
        Case "boolHashFalse": f = "#FALSE#": v = f
        Case "llFromDec": ll = CDec("9007199254740993"): v = CStr(ll)
        Case "curFromDec": c = CDec("922337203685477.5807"): v = CStr(c)
        Case "curTooBig": c = 922337203685477.625#: v = c
        Case "ccurNearMax": v = CCur(922337203685477.5#)
        Case "curLit": v = CStr(0.00005@)
        Case "setObjInt": vv = 5: Set o = vv: v = TypeName(o)
        Case "setVarInt": vv = 5: Set w = vv: v = TypeName(w)
        Case "memberOnInt": vv = 5: v = vv.Foo
        Case "infText": v = Special(&HF0, &H7F)
        Case "ninfText": v = Special(&HF0, &HFF)
        Case "nanText": v = Special(&HF8, &HFF)
        Case Else: v = "unknown case"
    End Select
    T = k & "=" & Sh(v): Exit Function
fail:
    T = k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim keys As Variant, i As Long, s As String
    keys = Split("byTrue cbyTrue dtBig dtSmall cdtBig strErr bytesToStr varStrToBytes dbl1d3 dblSpSign dblThou dblDollar dblParen dblTrailMinus dblE400 dblHex dblPad llBigStr boolHashTrue boolHashFalse llFromDec curFromDec curTooBig ccurNearMax curLit setObjInt setVarInt memberOnInt infText ninfText nanText", " ")
    For i = 0 To UBound(keys)
        s = s & T(keys(i)) & vbLf
    Next
    Run = s
End Function
