Function Rz(e As Long, val As Variant) As String
  If e <> 0 Then
    Rz = "ERR" & e
  ElseIf IsNull(val) Then
    Rz = "Null"
  ElseIf VarType(val) = vbDate Then
    Rz = "Date:" & Trim$(Str$(CDbl(val)))
  Else
    Rz = TypeName(val) & ":" & CStr(val)
  End If
End Function
Function L000() As String
  Dim r As Variant, out As String, vA0 As Variant, vA1 As Variant, vA2 As Variant, vA3 As Variant, vA4 As Variant, vA5 As Variant, vA6 As Variant, vA7 As Variant, vA8 As Variant, vA9 As Variant, vA10 As Variant, vA11 As Variant, vD0 As Variant, vD1 As Variant, vD2 As Variant, vD3 As Variant, vD4 As Variant, vD5 As Variant, vD6 As Variant, vD7 As Variant, vD8 As Variant, vD9 As Variant, vD10 As Variant, vD11 As Variant, vD12 As Variant, vD13 As Variant, vD14 As Variant, vD15 As Variant, vD16 As Variant, vD17 As Variant
  On Error Resume Next
  vA0 = "": vA1 = "a": vA2 = "abc": vA3 = "Hello World": vA4 = "  pad  ": vA5 = "12": vA6 = "-3.5": vA7 = "1e3": vA8 = "&H1F": vA9 = "abc abc": vA10 = Null: vA11 = Empty
  vD0 = 0: vD1 = 1: vD2 = -1: vD3 = 0.5: vD4 = -0.5: vD5 = 1.5: vD6 = 2.5: vD7 = -1.5: vD8 = -2.5: vD9 = 3.5: vD10 = 100.5: vD11 = 255.5: vD12 = 32767.5: vD13 = 2147483647.5: vD14 = 1E+10: vD15 = 1E+20: vD16 = 123456.789: vD17 = -123456.789
  Err.Clear: r = Empty: r = Len(vA0): out = out & "Len.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Len(vA1): out = out & "Len.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Len(vA2): out = out & "Len.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Len(vA3): out = out & "Len.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Len(vA4): out = out & "Len.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Len(vA5): out = out & "Len.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Len(vA6): out = out & "Len.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Len(vA7): out = out & "Len.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Len(vA8): out = out & "Len.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Len(vA9): out = out & "Len.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Len(vA10): out = out & "Len.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Len(vA11): out = out & "Len.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = UCase(vA0): out = out & "UCase.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = LCase(vA0): out = out & "LCase.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Trim(vA0): out = out & "Trim.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = LTrim(vA0): out = out & "LTrim.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = RTrim(vA0): out = out & "RTrim.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Asc(vA0): out = out & "Asc.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrReverse(vA0): out = out & "StrReverse.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = UCase(vA1): out = out & "UCase.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = LCase(vA1): out = out & "LCase.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Trim(vA1): out = out & "Trim.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = LTrim(vA1): out = out & "LTrim.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = RTrim(vA1): out = out & "RTrim.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Asc(vA1): out = out & "Asc.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrReverse(vA1): out = out & "StrReverse.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = UCase(vA2): out = out & "UCase.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = LCase(vA2): out = out & "LCase.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Trim(vA2): out = out & "Trim.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = LTrim(vA2): out = out & "LTrim.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = RTrim(vA2): out = out & "RTrim.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Asc(vA2): out = out & "Asc.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrReverse(vA2): out = out & "StrReverse.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = UCase(vA3): out = out & "UCase.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = LCase(vA3): out = out & "LCase.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Trim(vA3): out = out & "Trim.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = LTrim(vA3): out = out & "LTrim.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = RTrim(vA3): out = out & "RTrim.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Asc(vA3): out = out & "Asc.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrReverse(vA3): out = out & "StrReverse.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = UCase(vA4): out = out & "UCase.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = LCase(vA4): out = out & "LCase.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Trim(vA4): out = out & "Trim.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = LTrim(vA4): out = out & "LTrim.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = RTrim(vA4): out = out & "RTrim.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Asc(vA4): out = out & "Asc.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrReverse(vA4): out = out & "StrReverse.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = UCase(vA5): out = out & "UCase.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = LCase(vA5): out = out & "LCase.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Trim(vA5): out = out & "Trim.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = LTrim(vA5): out = out & "LTrim.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = RTrim(vA5): out = out & "RTrim.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Asc(vA5): out = out & "Asc.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrReverse(vA5): out = out & "StrReverse.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = UCase(vA6): out = out & "UCase.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = LCase(vA6): out = out & "LCase.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Trim(vA6): out = out & "Trim.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = LTrim(vA6): out = out & "LTrim.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = RTrim(vA6): out = out & "RTrim.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Asc(vA6): out = out & "Asc.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrReverse(vA6): out = out & "StrReverse.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = UCase(vA7): out = out & "UCase.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = LCase(vA7): out = out & "LCase.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Trim(vA7): out = out & "Trim.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = LTrim(vA7): out = out & "LTrim.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = RTrim(vA7): out = out & "RTrim.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Asc(vA7): out = out & "Asc.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrReverse(vA7): out = out & "StrReverse.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = UCase(vA8): out = out & "UCase.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = LCase(vA8): out = out & "LCase.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Trim(vA8): out = out & "Trim.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = LTrim(vA8): out = out & "LTrim.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = RTrim(vA8): out = out & "RTrim.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Asc(vA8): out = out & "Asc.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrReverse(vA8): out = out & "StrReverse.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = UCase(vA9): out = out & "UCase.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = LCase(vA9): out = out & "LCase.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Trim(vA9): out = out & "Trim.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = LTrim(vA9): out = out & "LTrim.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = RTrim(vA9): out = out & "RTrim.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Asc(vA9): out = out & "Asc.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrReverse(vA9): out = out & "StrReverse.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = UCase(vA10): out = out & "UCase.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = LCase(vA10): out = out & "LCase.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Trim(vA10): out = out & "Trim.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = LTrim(vA10): out = out & "LTrim.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = RTrim(vA10): out = out & "RTrim.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Asc(vA10): out = out & "Asc.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrReverse(vA10): out = out & "StrReverse.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = UCase(vA11): out = out & "UCase.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = LCase(vA11): out = out & "LCase.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Trim(vA11): out = out & "Trim.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = LTrim(vA11): out = out & "LTrim.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = RTrim(vA11): out = out & "RTrim.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Asc(vA11): out = out & "Asc.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrReverse(vA11): out = out & "StrReverse.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA0, -1): out = out & "Left.0.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA0, -1): out = out & "Right.0.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA0, -1): out = out & "Mid.0.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA0, -1, 2): out = out & "Mid3.0.-1=" & Rz(Err.Number, r) & vbLf
  L000 = out
End Function
Function L001() As String
  Dim r As Variant, out As String, vA0 As Variant, vA1 As Variant, vA2 As Variant, vA3 As Variant, vA4 As Variant, vA5 As Variant, vA6 As Variant, vA7 As Variant, vA8 As Variant, vA9 As Variant, vA10 As Variant, vA11 As Variant, vD0 As Variant, vD1 As Variant, vD2 As Variant, vD3 As Variant, vD4 As Variant, vD5 As Variant, vD6 As Variant, vD7 As Variant, vD8 As Variant, vD9 As Variant, vD10 As Variant, vD11 As Variant, vD12 As Variant, vD13 As Variant, vD14 As Variant, vD15 As Variant, vD16 As Variant, vD17 As Variant
  On Error Resume Next
  vA0 = "": vA1 = "a": vA2 = "abc": vA3 = "Hello World": vA4 = "  pad  ": vA5 = "12": vA6 = "-3.5": vA7 = "1e3": vA8 = "&H1F": vA9 = "abc abc": vA10 = Null: vA11 = Empty
  vD0 = 0: vD1 = 1: vD2 = -1: vD3 = 0.5: vD4 = -0.5: vD5 = 1.5: vD6 = 2.5: vD7 = -1.5: vD8 = -2.5: vD9 = 3.5: vD10 = 100.5: vD11 = 255.5: vD12 = 32767.5: vD13 = 2147483647.5: vD14 = 1E+10: vD15 = 1E+20: vD16 = 123456.789: vD17 = -123456.789
  Err.Clear: r = Empty: r = Left(vA0, 0): out = out & "Left.0.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA0, 0): out = out & "Right.0.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA0, 0): out = out & "Mid.0.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA0, 0, 2): out = out & "Mid3.0.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA0, 1): out = out & "Left.0.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA0, 1): out = out & "Right.0.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA0, 1): out = out & "Mid.0.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA0, 1, 2): out = out & "Mid3.0.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA0, 2): out = out & "Left.0.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA0, 2): out = out & "Right.0.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA0, 2): out = out & "Mid.0.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA0, 2, 2): out = out & "Mid3.0.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA0, 5): out = out & "Left.0.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA0, 5): out = out & "Right.0.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA0, 5): out = out & "Mid.0.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA0, 5, 2): out = out & "Mid3.0.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA0, 100): out = out & "Left.0.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA0, 100): out = out & "Right.0.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA0, 100): out = out & "Mid.0.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA0, 100, 2): out = out & "Mid3.0.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA1, -1): out = out & "Left.1.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA1, -1): out = out & "Right.1.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA1, -1): out = out & "Mid.1.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA1, -1, 2): out = out & "Mid3.1.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA1, 0): out = out & "Left.1.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA1, 0): out = out & "Right.1.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA1, 0): out = out & "Mid.1.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA1, 0, 2): out = out & "Mid3.1.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA1, 1): out = out & "Left.1.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA1, 1): out = out & "Right.1.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA1, 1): out = out & "Mid.1.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA1, 1, 2): out = out & "Mid3.1.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA1, 2): out = out & "Left.1.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA1, 2): out = out & "Right.1.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA1, 2): out = out & "Mid.1.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA1, 2, 2): out = out & "Mid3.1.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA1, 5): out = out & "Left.1.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA1, 5): out = out & "Right.1.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA1, 5): out = out & "Mid.1.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA1, 5, 2): out = out & "Mid3.1.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA1, 100): out = out & "Left.1.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA1, 100): out = out & "Right.1.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA1, 100): out = out & "Mid.1.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA1, 100, 2): out = out & "Mid3.1.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA2, -1): out = out & "Left.2.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA2, -1): out = out & "Right.2.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA2, -1): out = out & "Mid.2.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA2, -1, 2): out = out & "Mid3.2.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA2, 0): out = out & "Left.2.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA2, 0): out = out & "Right.2.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA2, 0): out = out & "Mid.2.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA2, 0, 2): out = out & "Mid3.2.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA2, 1): out = out & "Left.2.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA2, 1): out = out & "Right.2.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA2, 1): out = out & "Mid.2.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA2, 1, 2): out = out & "Mid3.2.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA2, 2): out = out & "Left.2.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA2, 2): out = out & "Right.2.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA2, 2): out = out & "Mid.2.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA2, 2, 2): out = out & "Mid3.2.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA2, 5): out = out & "Left.2.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA2, 5): out = out & "Right.2.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA2, 5): out = out & "Mid.2.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA2, 5, 2): out = out & "Mid3.2.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA2, 100): out = out & "Left.2.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA2, 100): out = out & "Right.2.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA2, 100): out = out & "Mid.2.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA2, 100, 2): out = out & "Mid3.2.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA3, -1): out = out & "Left.3.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA3, -1): out = out & "Right.3.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA3, -1): out = out & "Mid.3.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA3, -1, 2): out = out & "Mid3.3.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA3, 0): out = out & "Left.3.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA3, 0): out = out & "Right.3.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA3, 0): out = out & "Mid.3.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA3, 0, 2): out = out & "Mid3.3.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA3, 1): out = out & "Left.3.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA3, 1): out = out & "Right.3.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA3, 1): out = out & "Mid.3.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA3, 1, 2): out = out & "Mid3.3.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA3, 2): out = out & "Left.3.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA3, 2): out = out & "Right.3.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA3, 2): out = out & "Mid.3.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA3, 2, 2): out = out & "Mid3.3.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA3, 5): out = out & "Left.3.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA3, 5): out = out & "Right.3.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA3, 5): out = out & "Mid.3.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA3, 5, 2): out = out & "Mid3.3.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA3, 100): out = out & "Left.3.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA3, 100): out = out & "Right.3.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA3, 100): out = out & "Mid.3.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA3, 100, 2): out = out & "Mid3.3.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA4, -1): out = out & "Left.4.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA4, -1): out = out & "Right.4.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA4, -1): out = out & "Mid.4.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA4, -1, 2): out = out & "Mid3.4.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA4, 0): out = out & "Left.4.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA4, 0): out = out & "Right.4.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA4, 0): out = out & "Mid.4.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA4, 0, 2): out = out & "Mid3.4.0=" & Rz(Err.Number, r) & vbLf
  L001 = out
End Function
Function L002() As String
  Dim r As Variant, out As String, vA0 As Variant, vA1 As Variant, vA2 As Variant, vA3 As Variant, vA4 As Variant, vA5 As Variant, vA6 As Variant, vA7 As Variant, vA8 As Variant, vA9 As Variant, vA10 As Variant, vA11 As Variant, vD0 As Variant, vD1 As Variant, vD2 As Variant, vD3 As Variant, vD4 As Variant, vD5 As Variant, vD6 As Variant, vD7 As Variant, vD8 As Variant, vD9 As Variant, vD10 As Variant, vD11 As Variant, vD12 As Variant, vD13 As Variant, vD14 As Variant, vD15 As Variant, vD16 As Variant, vD17 As Variant
  On Error Resume Next
  vA0 = "": vA1 = "a": vA2 = "abc": vA3 = "Hello World": vA4 = "  pad  ": vA5 = "12": vA6 = "-3.5": vA7 = "1e3": vA8 = "&H1F": vA9 = "abc abc": vA10 = Null: vA11 = Empty
  vD0 = 0: vD1 = 1: vD2 = -1: vD3 = 0.5: vD4 = -0.5: vD5 = 1.5: vD6 = 2.5: vD7 = -1.5: vD8 = -2.5: vD9 = 3.5: vD10 = 100.5: vD11 = 255.5: vD12 = 32767.5: vD13 = 2147483647.5: vD14 = 1E+10: vD15 = 1E+20: vD16 = 123456.789: vD17 = -123456.789
  Err.Clear: r = Empty: r = Left(vA4, 1): out = out & "Left.4.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA4, 1): out = out & "Right.4.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA4, 1): out = out & "Mid.4.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA4, 1, 2): out = out & "Mid3.4.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA4, 2): out = out & "Left.4.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA4, 2): out = out & "Right.4.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA4, 2): out = out & "Mid.4.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA4, 2, 2): out = out & "Mid3.4.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA4, 5): out = out & "Left.4.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA4, 5): out = out & "Right.4.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA4, 5): out = out & "Mid.4.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA4, 5, 2): out = out & "Mid3.4.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA4, 100): out = out & "Left.4.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA4, 100): out = out & "Right.4.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA4, 100): out = out & "Mid.4.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA4, 100, 2): out = out & "Mid3.4.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA5, -1): out = out & "Left.5.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA5, -1): out = out & "Right.5.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA5, -1): out = out & "Mid.5.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA5, -1, 2): out = out & "Mid3.5.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA5, 0): out = out & "Left.5.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA5, 0): out = out & "Right.5.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA5, 0): out = out & "Mid.5.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA5, 0, 2): out = out & "Mid3.5.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA5, 1): out = out & "Left.5.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA5, 1): out = out & "Right.5.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA5, 1): out = out & "Mid.5.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA5, 1, 2): out = out & "Mid3.5.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA5, 2): out = out & "Left.5.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA5, 2): out = out & "Right.5.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA5, 2): out = out & "Mid.5.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA5, 2, 2): out = out & "Mid3.5.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA5, 5): out = out & "Left.5.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA5, 5): out = out & "Right.5.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA5, 5): out = out & "Mid.5.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA5, 5, 2): out = out & "Mid3.5.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA5, 100): out = out & "Left.5.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA5, 100): out = out & "Right.5.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA5, 100): out = out & "Mid.5.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA5, 100, 2): out = out & "Mid3.5.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA6, -1): out = out & "Left.6.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA6, -1): out = out & "Right.6.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA6, -1): out = out & "Mid.6.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA6, -1, 2): out = out & "Mid3.6.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA6, 0): out = out & "Left.6.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA6, 0): out = out & "Right.6.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA6, 0): out = out & "Mid.6.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA6, 0, 2): out = out & "Mid3.6.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA6, 1): out = out & "Left.6.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA6, 1): out = out & "Right.6.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA6, 1): out = out & "Mid.6.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA6, 1, 2): out = out & "Mid3.6.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA6, 2): out = out & "Left.6.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA6, 2): out = out & "Right.6.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA6, 2): out = out & "Mid.6.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA6, 2, 2): out = out & "Mid3.6.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA6, 5): out = out & "Left.6.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA6, 5): out = out & "Right.6.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA6, 5): out = out & "Mid.6.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA6, 5, 2): out = out & "Mid3.6.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA6, 100): out = out & "Left.6.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA6, 100): out = out & "Right.6.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA6, 100): out = out & "Mid.6.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA6, 100, 2): out = out & "Mid3.6.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA7, -1): out = out & "Left.7.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA7, -1): out = out & "Right.7.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA7, -1): out = out & "Mid.7.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA7, -1, 2): out = out & "Mid3.7.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA7, 0): out = out & "Left.7.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA7, 0): out = out & "Right.7.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA7, 0): out = out & "Mid.7.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA7, 0, 2): out = out & "Mid3.7.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA7, 1): out = out & "Left.7.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA7, 1): out = out & "Right.7.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA7, 1): out = out & "Mid.7.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA7, 1, 2): out = out & "Mid3.7.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA7, 2): out = out & "Left.7.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA7, 2): out = out & "Right.7.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA7, 2): out = out & "Mid.7.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA7, 2, 2): out = out & "Mid3.7.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA7, 5): out = out & "Left.7.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA7, 5): out = out & "Right.7.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA7, 5): out = out & "Mid.7.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA7, 5, 2): out = out & "Mid3.7.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA7, 100): out = out & "Left.7.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA7, 100): out = out & "Right.7.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA7, 100): out = out & "Mid.7.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA7, 100, 2): out = out & "Mid3.7.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA8, -1): out = out & "Left.8.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA8, -1): out = out & "Right.8.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA8, -1): out = out & "Mid.8.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA8, -1, 2): out = out & "Mid3.8.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA8, 0): out = out & "Left.8.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA8, 0): out = out & "Right.8.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA8, 0): out = out & "Mid.8.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA8, 0, 2): out = out & "Mid3.8.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA8, 1): out = out & "Left.8.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA8, 1): out = out & "Right.8.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA8, 1): out = out & "Mid.8.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA8, 1, 2): out = out & "Mid3.8.1=" & Rz(Err.Number, r) & vbLf
  L002 = out
End Function
Function L003() As String
  Dim r As Variant, out As String, vA0 As Variant, vA1 As Variant, vA2 As Variant, vA3 As Variant, vA4 As Variant, vA5 As Variant, vA6 As Variant, vA7 As Variant, vA8 As Variant, vA9 As Variant, vA10 As Variant, vA11 As Variant, vD0 As Variant, vD1 As Variant, vD2 As Variant, vD3 As Variant, vD4 As Variant, vD5 As Variant, vD6 As Variant, vD7 As Variant, vD8 As Variant, vD9 As Variant, vD10 As Variant, vD11 As Variant, vD12 As Variant, vD13 As Variant, vD14 As Variant, vD15 As Variant, vD16 As Variant, vD17 As Variant
  On Error Resume Next
  vA0 = "": vA1 = "a": vA2 = "abc": vA3 = "Hello World": vA4 = "  pad  ": vA5 = "12": vA6 = "-3.5": vA7 = "1e3": vA8 = "&H1F": vA9 = "abc abc": vA10 = Null: vA11 = Empty
  vD0 = 0: vD1 = 1: vD2 = -1: vD3 = 0.5: vD4 = -0.5: vD5 = 1.5: vD6 = 2.5: vD7 = -1.5: vD8 = -2.5: vD9 = 3.5: vD10 = 100.5: vD11 = 255.5: vD12 = 32767.5: vD13 = 2147483647.5: vD14 = 1E+10: vD15 = 1E+20: vD16 = 123456.789: vD17 = -123456.789
  Err.Clear: r = Empty: r = Left(vA8, 2): out = out & "Left.8.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA8, 2): out = out & "Right.8.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA8, 2): out = out & "Mid.8.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA8, 2, 2): out = out & "Mid3.8.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA8, 5): out = out & "Left.8.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA8, 5): out = out & "Right.8.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA8, 5): out = out & "Mid.8.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA8, 5, 2): out = out & "Mid3.8.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA8, 100): out = out & "Left.8.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA8, 100): out = out & "Right.8.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA8, 100): out = out & "Mid.8.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA8, 100, 2): out = out & "Mid3.8.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA9, -1): out = out & "Left.9.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA9, -1): out = out & "Right.9.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA9, -1): out = out & "Mid.9.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA9, -1, 2): out = out & "Mid3.9.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA9, 0): out = out & "Left.9.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA9, 0): out = out & "Right.9.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA9, 0): out = out & "Mid.9.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA9, 0, 2): out = out & "Mid3.9.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA9, 1): out = out & "Left.9.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA9, 1): out = out & "Right.9.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA9, 1): out = out & "Mid.9.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA9, 1, 2): out = out & "Mid3.9.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA9, 2): out = out & "Left.9.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA9, 2): out = out & "Right.9.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA9, 2): out = out & "Mid.9.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA9, 2, 2): out = out & "Mid3.9.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA9, 5): out = out & "Left.9.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA9, 5): out = out & "Right.9.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA9, 5): out = out & "Mid.9.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA9, 5, 2): out = out & "Mid3.9.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA9, 100): out = out & "Left.9.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA9, 100): out = out & "Right.9.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA9, 100): out = out & "Mid.9.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA9, 100, 2): out = out & "Mid3.9.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA10, -1): out = out & "Left.10.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA10, -1): out = out & "Right.10.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA10, -1): out = out & "Mid.10.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA10, -1, 2): out = out & "Mid3.10.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA10, 0): out = out & "Left.10.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA10, 0): out = out & "Right.10.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA10, 0): out = out & "Mid.10.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA10, 0, 2): out = out & "Mid3.10.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA10, 1): out = out & "Left.10.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA10, 1): out = out & "Right.10.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA10, 1): out = out & "Mid.10.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA10, 1, 2): out = out & "Mid3.10.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA10, 2): out = out & "Left.10.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA10, 2): out = out & "Right.10.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA10, 2): out = out & "Mid.10.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA10, 2, 2): out = out & "Mid3.10.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA10, 5): out = out & "Left.10.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA10, 5): out = out & "Right.10.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA10, 5): out = out & "Mid.10.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA10, 5, 2): out = out & "Mid3.10.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA10, 100): out = out & "Left.10.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA10, 100): out = out & "Right.10.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA10, 100): out = out & "Mid.10.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA10, 100, 2): out = out & "Mid3.10.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA11, -1): out = out & "Left.11.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA11, -1): out = out & "Right.11.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA11, -1): out = out & "Mid.11.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA11, -1, 2): out = out & "Mid3.11.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA11, 0): out = out & "Left.11.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA11, 0): out = out & "Right.11.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA11, 0): out = out & "Mid.11.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA11, 0, 2): out = out & "Mid3.11.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA11, 1): out = out & "Left.11.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA11, 1): out = out & "Right.11.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA11, 1): out = out & "Mid.11.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA11, 1, 2): out = out & "Mid3.11.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA11, 2): out = out & "Left.11.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA11, 2): out = out & "Right.11.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA11, 2): out = out & "Mid.11.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA11, 2, 2): out = out & "Mid3.11.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA11, 5): out = out & "Left.11.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA11, 5): out = out & "Right.11.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA11, 5): out = out & "Mid.11.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA11, 5, 2): out = out & "Mid3.11.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Left(vA11, 100): out = out & "Left.11.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Right(vA11, 100): out = out & "Right.11.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA11, 100): out = out & "Mid.11.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Mid(vA11, 100, 2): out = out & "Mid3.11.100=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(-1, vA2, vA0): out = out & "InStr.2.0.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(0, vA2, vA0): out = out & "InStr.2.0.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(1, vA2, vA0): out = out & "InStr.2.0.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(3, vA2, vA0): out = out & "InStr.2.0.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(20, vA2, vA0): out = out & "InStr.2.0.20=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStrRev(vA2, vA0): out = out & "InStrRev.2.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA2, vA0): out = out & "StrComp.2.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA2, vA0, vbTextCompare): out = out & "StrCompT.2.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Replace(vA2, vA0, "X"): out = out & "Replace.2.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(-1, vA2, vA1): out = out & "InStr.2.1.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(0, vA2, vA1): out = out & "InStr.2.1.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(1, vA2, vA1): out = out & "InStr.2.1.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(3, vA2, vA1): out = out & "InStr.2.1.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(20, vA2, vA1): out = out & "InStr.2.1.20=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStrRev(vA2, vA1): out = out & "InStrRev.2.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA2, vA1): out = out & "StrComp.2.1=" & Rz(Err.Number, r) & vbLf
  L003 = out
End Function
Function L004() As String
  Dim r As Variant, out As String, vA0 As Variant, vA1 As Variant, vA2 As Variant, vA3 As Variant, vA4 As Variant, vA5 As Variant, vA6 As Variant, vA7 As Variant, vA8 As Variant, vA9 As Variant, vA10 As Variant, vA11 As Variant, vD0 As Variant, vD1 As Variant, vD2 As Variant, vD3 As Variant, vD4 As Variant, vD5 As Variant, vD6 As Variant, vD7 As Variant, vD8 As Variant, vD9 As Variant, vD10 As Variant, vD11 As Variant, vD12 As Variant, vD13 As Variant, vD14 As Variant, vD15 As Variant, vD16 As Variant, vD17 As Variant
  On Error Resume Next
  vA0 = "": vA1 = "a": vA2 = "abc": vA3 = "Hello World": vA4 = "  pad  ": vA5 = "12": vA6 = "-3.5": vA7 = "1e3": vA8 = "&H1F": vA9 = "abc abc": vA10 = Null: vA11 = Empty
  vD0 = 0: vD1 = 1: vD2 = -1: vD3 = 0.5: vD4 = -0.5: vD5 = 1.5: vD6 = 2.5: vD7 = -1.5: vD8 = -2.5: vD9 = 3.5: vD10 = 100.5: vD11 = 255.5: vD12 = 32767.5: vD13 = 2147483647.5: vD14 = 1E+10: vD15 = 1E+20: vD16 = 123456.789: vD17 = -123456.789
  Err.Clear: r = Empty: r = StrComp(vA2, vA1, vbTextCompare): out = out & "StrCompT.2.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Replace(vA2, vA1, "X"): out = out & "Replace.2.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(-1, vA2, vA2): out = out & "InStr.2.2.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(0, vA2, vA2): out = out & "InStr.2.2.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(1, vA2, vA2): out = out & "InStr.2.2.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(3, vA2, vA2): out = out & "InStr.2.2.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(20, vA2, vA2): out = out & "InStr.2.2.20=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStrRev(vA2, vA2): out = out & "InStrRev.2.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA2, vA2): out = out & "StrComp.2.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA2, vA2, vbTextCompare): out = out & "StrCompT.2.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Replace(vA2, vA2, "X"): out = out & "Replace.2.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(-1, vA2, vA3): out = out & "InStr.2.3.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(0, vA2, vA3): out = out & "InStr.2.3.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(1, vA2, vA3): out = out & "InStr.2.3.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(3, vA2, vA3): out = out & "InStr.2.3.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(20, vA2, vA3): out = out & "InStr.2.3.20=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStrRev(vA2, vA3): out = out & "InStrRev.2.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA2, vA3): out = out & "StrComp.2.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA2, vA3, vbTextCompare): out = out & "StrCompT.2.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Replace(vA2, vA3, "X"): out = out & "Replace.2.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(-1, vA2, vA9): out = out & "InStr.2.9.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(0, vA2, vA9): out = out & "InStr.2.9.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(1, vA2, vA9): out = out & "InStr.2.9.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(3, vA2, vA9): out = out & "InStr.2.9.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(20, vA2, vA9): out = out & "InStr.2.9.20=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStrRev(vA2, vA9): out = out & "InStrRev.2.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA2, vA9): out = out & "StrComp.2.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA2, vA9, vbTextCompare): out = out & "StrCompT.2.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Replace(vA2, vA9, "X"): out = out & "Replace.2.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(-1, vA3, vA0): out = out & "InStr.3.0.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(0, vA3, vA0): out = out & "InStr.3.0.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(1, vA3, vA0): out = out & "InStr.3.0.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(3, vA3, vA0): out = out & "InStr.3.0.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(20, vA3, vA0): out = out & "InStr.3.0.20=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStrRev(vA3, vA0): out = out & "InStrRev.3.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA3, vA0): out = out & "StrComp.3.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA3, vA0, vbTextCompare): out = out & "StrCompT.3.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Replace(vA3, vA0, "X"): out = out & "Replace.3.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(-1, vA3, vA1): out = out & "InStr.3.1.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(0, vA3, vA1): out = out & "InStr.3.1.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(1, vA3, vA1): out = out & "InStr.3.1.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(3, vA3, vA1): out = out & "InStr.3.1.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(20, vA3, vA1): out = out & "InStr.3.1.20=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStrRev(vA3, vA1): out = out & "InStrRev.3.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA3, vA1): out = out & "StrComp.3.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA3, vA1, vbTextCompare): out = out & "StrCompT.3.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Replace(vA3, vA1, "X"): out = out & "Replace.3.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(-1, vA3, vA2): out = out & "InStr.3.2.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(0, vA3, vA2): out = out & "InStr.3.2.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(1, vA3, vA2): out = out & "InStr.3.2.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(3, vA3, vA2): out = out & "InStr.3.2.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(20, vA3, vA2): out = out & "InStr.3.2.20=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStrRev(vA3, vA2): out = out & "InStrRev.3.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA3, vA2): out = out & "StrComp.3.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA3, vA2, vbTextCompare): out = out & "StrCompT.3.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Replace(vA3, vA2, "X"): out = out & "Replace.3.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(-1, vA3, vA3): out = out & "InStr.3.3.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(0, vA3, vA3): out = out & "InStr.3.3.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(1, vA3, vA3): out = out & "InStr.3.3.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(3, vA3, vA3): out = out & "InStr.3.3.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(20, vA3, vA3): out = out & "InStr.3.3.20=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStrRev(vA3, vA3): out = out & "InStrRev.3.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA3, vA3): out = out & "StrComp.3.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA3, vA3, vbTextCompare): out = out & "StrCompT.3.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Replace(vA3, vA3, "X"): out = out & "Replace.3.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(-1, vA3, vA9): out = out & "InStr.3.9.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(0, vA3, vA9): out = out & "InStr.3.9.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(1, vA3, vA9): out = out & "InStr.3.9.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(3, vA3, vA9): out = out & "InStr.3.9.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(20, vA3, vA9): out = out & "InStr.3.9.20=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStrRev(vA3, vA9): out = out & "InStrRev.3.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA3, vA9): out = out & "StrComp.3.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA3, vA9, vbTextCompare): out = out & "StrCompT.3.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Replace(vA3, vA9, "X"): out = out & "Replace.3.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(-1, vA7, vA0): out = out & "InStr.7.0.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(0, vA7, vA0): out = out & "InStr.7.0.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(1, vA7, vA0): out = out & "InStr.7.0.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(3, vA7, vA0): out = out & "InStr.7.0.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(20, vA7, vA0): out = out & "InStr.7.0.20=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStrRev(vA7, vA0): out = out & "InStrRev.7.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA7, vA0): out = out & "StrComp.7.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA7, vA0, vbTextCompare): out = out & "StrCompT.7.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Replace(vA7, vA0, "X"): out = out & "Replace.7.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(-1, vA7, vA1): out = out & "InStr.7.1.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(0, vA7, vA1): out = out & "InStr.7.1.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(1, vA7, vA1): out = out & "InStr.7.1.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(3, vA7, vA1): out = out & "InStr.7.1.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(20, vA7, vA1): out = out & "InStr.7.1.20=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStrRev(vA7, vA1): out = out & "InStrRev.7.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA7, vA1): out = out & "StrComp.7.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA7, vA1, vbTextCompare): out = out & "StrCompT.7.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Replace(vA7, vA1, "X"): out = out & "Replace.7.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(-1, vA7, vA2): out = out & "InStr.7.2.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(0, vA7, vA2): out = out & "InStr.7.2.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(1, vA7, vA2): out = out & "InStr.7.2.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(3, vA7, vA2): out = out & "InStr.7.2.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(20, vA7, vA2): out = out & "InStr.7.2.20=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStrRev(vA7, vA2): out = out & "InStrRev.7.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA7, vA2): out = out & "StrComp.7.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA7, vA2, vbTextCompare): out = out & "StrCompT.7.2=" & Rz(Err.Number, r) & vbLf
  L004 = out
End Function
Function L005() As String
  Dim r As Variant, out As String, vA0 As Variant, vA1 As Variant, vA2 As Variant, vA3 As Variant, vA4 As Variant, vA5 As Variant, vA6 As Variant, vA7 As Variant, vA8 As Variant, vA9 As Variant, vA10 As Variant, vA11 As Variant, vD0 As Variant, vD1 As Variant, vD2 As Variant, vD3 As Variant, vD4 As Variant, vD5 As Variant, vD6 As Variant, vD7 As Variant, vD8 As Variant, vD9 As Variant, vD10 As Variant, vD11 As Variant, vD12 As Variant, vD13 As Variant, vD14 As Variant, vD15 As Variant, vD16 As Variant, vD17 As Variant
  On Error Resume Next
  vA0 = "": vA1 = "a": vA2 = "abc": vA3 = "Hello World": vA4 = "  pad  ": vA5 = "12": vA6 = "-3.5": vA7 = "1e3": vA8 = "&H1F": vA9 = "abc abc": vA10 = Null: vA11 = Empty
  vD0 = 0: vD1 = 1: vD2 = -1: vD3 = 0.5: vD4 = -0.5: vD5 = 1.5: vD6 = 2.5: vD7 = -1.5: vD8 = -2.5: vD9 = 3.5: vD10 = 100.5: vD11 = 255.5: vD12 = 32767.5: vD13 = 2147483647.5: vD14 = 1E+10: vD15 = 1E+20: vD16 = 123456.789: vD17 = -123456.789
  Err.Clear: r = Empty: r = Replace(vA7, vA2, "X"): out = out & "Replace.7.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(-1, vA7, vA3): out = out & "InStr.7.3.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(0, vA7, vA3): out = out & "InStr.7.3.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(1, vA7, vA3): out = out & "InStr.7.3.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(3, vA7, vA3): out = out & "InStr.7.3.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(20, vA7, vA3): out = out & "InStr.7.3.20=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStrRev(vA7, vA3): out = out & "InStrRev.7.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA7, vA3): out = out & "StrComp.7.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA7, vA3, vbTextCompare): out = out & "StrCompT.7.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Replace(vA7, vA3, "X"): out = out & "Replace.7.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(-1, vA7, vA9): out = out & "InStr.7.9.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(0, vA7, vA9): out = out & "InStr.7.9.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(1, vA7, vA9): out = out & "InStr.7.9.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(3, vA7, vA9): out = out & "InStr.7.9.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(20, vA7, vA9): out = out & "InStr.7.9.20=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStrRev(vA7, vA9): out = out & "InStrRev.7.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA7, vA9): out = out & "StrComp.7.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA7, vA9, vbTextCompare): out = out & "StrCompT.7.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Replace(vA7, vA9, "X"): out = out & "Replace.7.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(-1, vA9, vA0): out = out & "InStr.9.0.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(0, vA9, vA0): out = out & "InStr.9.0.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(1, vA9, vA0): out = out & "InStr.9.0.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(3, vA9, vA0): out = out & "InStr.9.0.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(20, vA9, vA0): out = out & "InStr.9.0.20=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStrRev(vA9, vA0): out = out & "InStrRev.9.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA9, vA0): out = out & "StrComp.9.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA9, vA0, vbTextCompare): out = out & "StrCompT.9.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Replace(vA9, vA0, "X"): out = out & "Replace.9.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(-1, vA9, vA1): out = out & "InStr.9.1.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(0, vA9, vA1): out = out & "InStr.9.1.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(1, vA9, vA1): out = out & "InStr.9.1.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(3, vA9, vA1): out = out & "InStr.9.1.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(20, vA9, vA1): out = out & "InStr.9.1.20=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStrRev(vA9, vA1): out = out & "InStrRev.9.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA9, vA1): out = out & "StrComp.9.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA9, vA1, vbTextCompare): out = out & "StrCompT.9.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Replace(vA9, vA1, "X"): out = out & "Replace.9.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(-1, vA9, vA2): out = out & "InStr.9.2.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(0, vA9, vA2): out = out & "InStr.9.2.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(1, vA9, vA2): out = out & "InStr.9.2.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(3, vA9, vA2): out = out & "InStr.9.2.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(20, vA9, vA2): out = out & "InStr.9.2.20=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStrRev(vA9, vA2): out = out & "InStrRev.9.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA9, vA2): out = out & "StrComp.9.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA9, vA2, vbTextCompare): out = out & "StrCompT.9.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Replace(vA9, vA2, "X"): out = out & "Replace.9.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(-1, vA9, vA3): out = out & "InStr.9.3.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(0, vA9, vA3): out = out & "InStr.9.3.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(1, vA9, vA3): out = out & "InStr.9.3.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(3, vA9, vA3): out = out & "InStr.9.3.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(20, vA9, vA3): out = out & "InStr.9.3.20=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStrRev(vA9, vA3): out = out & "InStrRev.9.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA9, vA3): out = out & "StrComp.9.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA9, vA3, vbTextCompare): out = out & "StrCompT.9.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Replace(vA9, vA3, "X"): out = out & "Replace.9.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(-1, vA9, vA9): out = out & "InStr.9.9.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(0, vA9, vA9): out = out & "InStr.9.9.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(1, vA9, vA9): out = out & "InStr.9.9.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(3, vA9, vA9): out = out & "InStr.9.9.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStr(20, vA9, vA9): out = out & "InStr.9.9.20=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = InStrRev(vA9, vA9): out = out & "InStrRev.9.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA9, vA9): out = out & "StrComp.9.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = StrComp(vA9, vA9, vbTextCompare): out = out & "StrCompT.9.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Replace(vA9, vA9, "X"): out = out & "Replace.9.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = String(-1, "ab"): out = out & "String.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Space(-1): out = out & "Space.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = String(0, "ab"): out = out & "String.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Space(0): out = out & "Space.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = String(1, "ab"): out = out & "String.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Space(1): out = out & "Space.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = String(3, "ab"): out = out & "String.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Space(3): out = out & "Space.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Chr(-1): out = out & "Chr.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = ChrW(-1): out = out & "ChrW.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Chr(0): out = out & "Chr.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = ChrW(0): out = out & "ChrW.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Chr(1): out = out & "Chr.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = ChrW(1): out = out & "ChrW.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Chr(65): out = out & "Chr.65=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = ChrW(65): out = out & "ChrW.65=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Chr(255): out = out & "Chr.255=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = ChrW(255): out = out & "ChrW.255=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Chr(256): out = out & "Chr.256=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = ChrW(256): out = out & "ChrW.256=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Chr(65535): out = out & "Chr.65535=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = ChrW(65535): out = out & "ChrW.65535=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Chr(65536): out = out & "Chr.65536=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = ChrW(65536): out = out & "ChrW.65536=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Int(vD0): out = out & "Int.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Fix(vD0): out = out & "Fix.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Abs(vD0): out = out & "Abs.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sgn(vD0): out = out & "Sgn.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vD0): out = out & "CInt.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vD0): out = out & "CLng.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vD0): out = out & "CByte.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vD0): out = out & "CDbl.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vD0): out = out & "CSng.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vD0): out = out & "CCur.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vD0): out = out & "CStr.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vD0): out = out & "CBool.0=" & Rz(Err.Number, r) & vbLf
  L005 = out
End Function
Function L006() As String
  Dim r As Variant, out As String, vA0 As Variant, vA1 As Variant, vA2 As Variant, vA3 As Variant, vA4 As Variant, vA5 As Variant, vA6 As Variant, vA7 As Variant, vA8 As Variant, vA9 As Variant, vA10 As Variant, vA11 As Variant, vD0 As Variant, vD1 As Variant, vD2 As Variant, vD3 As Variant, vD4 As Variant, vD5 As Variant, vD6 As Variant, vD7 As Variant, vD8 As Variant, vD9 As Variant, vD10 As Variant, vD11 As Variant, vD12 As Variant, vD13 As Variant, vD14 As Variant, vD15 As Variant, vD16 As Variant, vD17 As Variant
  On Error Resume Next
  vA0 = "": vA1 = "a": vA2 = "abc": vA3 = "Hello World": vA4 = "  pad  ": vA5 = "12": vA6 = "-3.5": vA7 = "1e3": vA8 = "&H1F": vA9 = "abc abc": vA10 = Null: vA11 = Empty
  vD0 = 0: vD1 = 1: vD2 = -1: vD3 = 0.5: vD4 = -0.5: vD5 = 1.5: vD6 = 2.5: vD7 = -1.5: vD8 = -2.5: vD9 = 3.5: vD10 = 100.5: vD11 = 255.5: vD12 = 32767.5: vD13 = 2147483647.5: vD14 = 1E+10: vD15 = 1E+20: vD16 = 123456.789: vD17 = -123456.789
  Err.Clear: r = Empty: r = Hex(vD0): out = out & "Hex.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Oct(vD0): out = out & "Oct.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sqr(vD0): out = out & "Sqr.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Exp(vD0): out = out & "Exp.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Log(vD0): out = out & "Log.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Atn(vD0): out = out & "Atn.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sin(vD0): out = out & "Sin.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Cos(vD0): out = out & "Cos.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Tan(vD0): out = out & "Tan.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD0, -1): out = out & "Round.0.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD0, 0): out = out & "Round.0.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD0, 1): out = out & "Round.0.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD0, 2): out = out & "Round.0.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD0, "0.00"): out = out & "Format.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD0, "General Number"): out = out & "FormatG.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD0, "#,##0.0"): out = out & "FormatP.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD0, "0.0E+00"): out = out & "FormatC.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD0, "0%"): out = out & "FormatPct.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(CStr(vD0)): out = out & "Val.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Int(vD1): out = out & "Int.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Fix(vD1): out = out & "Fix.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Abs(vD1): out = out & "Abs.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sgn(vD1): out = out & "Sgn.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vD1): out = out & "CInt.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vD1): out = out & "CLng.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vD1): out = out & "CByte.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vD1): out = out & "CDbl.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vD1): out = out & "CSng.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vD1): out = out & "CCur.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vD1): out = out & "CStr.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vD1): out = out & "CBool.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Hex(vD1): out = out & "Hex.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Oct(vD1): out = out & "Oct.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sqr(vD1): out = out & "Sqr.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Exp(vD1): out = out & "Exp.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Log(vD1): out = out & "Log.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Atn(vD1): out = out & "Atn.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sin(vD1): out = out & "Sin.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Cos(vD1): out = out & "Cos.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Tan(vD1): out = out & "Tan.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD1, -1): out = out & "Round.1.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD1, 0): out = out & "Round.1.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD1, 1): out = out & "Round.1.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD1, 2): out = out & "Round.1.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD1, "0.00"): out = out & "Format.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD1, "General Number"): out = out & "FormatG.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD1, "#,##0.0"): out = out & "FormatP.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD1, "0.0E+00"): out = out & "FormatC.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD1, "0%"): out = out & "FormatPct.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(CStr(vD1)): out = out & "Val.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Int(vD2): out = out & "Int.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Fix(vD2): out = out & "Fix.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Abs(vD2): out = out & "Abs.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sgn(vD2): out = out & "Sgn.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vD2): out = out & "CInt.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vD2): out = out & "CLng.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vD2): out = out & "CByte.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vD2): out = out & "CDbl.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vD2): out = out & "CSng.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vD2): out = out & "CCur.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vD2): out = out & "CStr.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vD2): out = out & "CBool.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Hex(vD2): out = out & "Hex.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Oct(vD2): out = out & "Oct.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sqr(vD2): out = out & "Sqr.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Exp(vD2): out = out & "Exp.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Log(vD2): out = out & "Log.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Atn(vD2): out = out & "Atn.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sin(vD2): out = out & "Sin.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Cos(vD2): out = out & "Cos.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Tan(vD2): out = out & "Tan.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD2, -1): out = out & "Round.2.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD2, 0): out = out & "Round.2.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD2, 1): out = out & "Round.2.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD2, 2): out = out & "Round.2.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD2, "0.00"): out = out & "Format.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD2, "General Number"): out = out & "FormatG.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD2, "#,##0.0"): out = out & "FormatP.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD2, "0.0E+00"): out = out & "FormatC.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD2, "0%"): out = out & "FormatPct.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(CStr(vD2)): out = out & "Val.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Int(vD3): out = out & "Int.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Fix(vD3): out = out & "Fix.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Abs(vD3): out = out & "Abs.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sgn(vD3): out = out & "Sgn.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vD3): out = out & "CInt.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vD3): out = out & "CLng.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vD3): out = out & "CByte.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vD3): out = out & "CDbl.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vD3): out = out & "CSng.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vD3): out = out & "CCur.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vD3): out = out & "CStr.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vD3): out = out & "CBool.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Hex(vD3): out = out & "Hex.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Oct(vD3): out = out & "Oct.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sqr(vD3): out = out & "Sqr.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Exp(vD3): out = out & "Exp.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Log(vD3): out = out & "Log.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Atn(vD3): out = out & "Atn.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sin(vD3): out = out & "Sin.3=" & Rz(Err.Number, r) & vbLf
  L006 = out
End Function
Function L007() As String
  Dim r As Variant, out As String, vA0 As Variant, vA1 As Variant, vA2 As Variant, vA3 As Variant, vA4 As Variant, vA5 As Variant, vA6 As Variant, vA7 As Variant, vA8 As Variant, vA9 As Variant, vA10 As Variant, vA11 As Variant, vD0 As Variant, vD1 As Variant, vD2 As Variant, vD3 As Variant, vD4 As Variant, vD5 As Variant, vD6 As Variant, vD7 As Variant, vD8 As Variant, vD9 As Variant, vD10 As Variant, vD11 As Variant, vD12 As Variant, vD13 As Variant, vD14 As Variant, vD15 As Variant, vD16 As Variant, vD17 As Variant
  On Error Resume Next
  vA0 = "": vA1 = "a": vA2 = "abc": vA3 = "Hello World": vA4 = "  pad  ": vA5 = "12": vA6 = "-3.5": vA7 = "1e3": vA8 = "&H1F": vA9 = "abc abc": vA10 = Null: vA11 = Empty
  vD0 = 0: vD1 = 1: vD2 = -1: vD3 = 0.5: vD4 = -0.5: vD5 = 1.5: vD6 = 2.5: vD7 = -1.5: vD8 = -2.5: vD9 = 3.5: vD10 = 100.5: vD11 = 255.5: vD12 = 32767.5: vD13 = 2147483647.5: vD14 = 1E+10: vD15 = 1E+20: vD16 = 123456.789: vD17 = -123456.789
  Err.Clear: r = Empty: r = Cos(vD3): out = out & "Cos.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Tan(vD3): out = out & "Tan.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD3, -1): out = out & "Round.3.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD3, 0): out = out & "Round.3.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD3, 1): out = out & "Round.3.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD3, 2): out = out & "Round.3.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD3, "0.00"): out = out & "Format.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD3, "General Number"): out = out & "FormatG.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD3, "#,##0.0"): out = out & "FormatP.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD3, "0.0E+00"): out = out & "FormatC.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD3, "0%"): out = out & "FormatPct.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(CStr(vD3)): out = out & "Val.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Int(vD4): out = out & "Int.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Fix(vD4): out = out & "Fix.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Abs(vD4): out = out & "Abs.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sgn(vD4): out = out & "Sgn.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vD4): out = out & "CInt.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vD4): out = out & "CLng.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vD4): out = out & "CByte.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vD4): out = out & "CDbl.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vD4): out = out & "CSng.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vD4): out = out & "CCur.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vD4): out = out & "CStr.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vD4): out = out & "CBool.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Hex(vD4): out = out & "Hex.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Oct(vD4): out = out & "Oct.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sqr(vD4): out = out & "Sqr.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Exp(vD4): out = out & "Exp.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Log(vD4): out = out & "Log.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Atn(vD4): out = out & "Atn.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sin(vD4): out = out & "Sin.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Cos(vD4): out = out & "Cos.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Tan(vD4): out = out & "Tan.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD4, -1): out = out & "Round.4.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD4, 0): out = out & "Round.4.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD4, 1): out = out & "Round.4.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD4, 2): out = out & "Round.4.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD4, "0.00"): out = out & "Format.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD4, "General Number"): out = out & "FormatG.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD4, "#,##0.0"): out = out & "FormatP.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD4, "0.0E+00"): out = out & "FormatC.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD4, "0%"): out = out & "FormatPct.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(CStr(vD4)): out = out & "Val.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Int(vD5): out = out & "Int.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Fix(vD5): out = out & "Fix.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Abs(vD5): out = out & "Abs.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sgn(vD5): out = out & "Sgn.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vD5): out = out & "CInt.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vD5): out = out & "CLng.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vD5): out = out & "CByte.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vD5): out = out & "CDbl.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vD5): out = out & "CSng.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vD5): out = out & "CCur.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vD5): out = out & "CStr.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vD5): out = out & "CBool.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Hex(vD5): out = out & "Hex.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Oct(vD5): out = out & "Oct.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sqr(vD5): out = out & "Sqr.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Exp(vD5): out = out & "Exp.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Log(vD5): out = out & "Log.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Atn(vD5): out = out & "Atn.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sin(vD5): out = out & "Sin.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Cos(vD5): out = out & "Cos.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Tan(vD5): out = out & "Tan.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD5, -1): out = out & "Round.5.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD5, 0): out = out & "Round.5.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD5, 1): out = out & "Round.5.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD5, 2): out = out & "Round.5.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD5, "0.00"): out = out & "Format.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD5, "General Number"): out = out & "FormatG.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD5, "#,##0.0"): out = out & "FormatP.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD5, "0.0E+00"): out = out & "FormatC.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD5, "0%"): out = out & "FormatPct.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(CStr(vD5)): out = out & "Val.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Int(vD6): out = out & "Int.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Fix(vD6): out = out & "Fix.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Abs(vD6): out = out & "Abs.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sgn(vD6): out = out & "Sgn.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vD6): out = out & "CInt.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vD6): out = out & "CLng.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vD6): out = out & "CByte.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vD6): out = out & "CDbl.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vD6): out = out & "CSng.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vD6): out = out & "CCur.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vD6): out = out & "CStr.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vD6): out = out & "CBool.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Hex(vD6): out = out & "Hex.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Oct(vD6): out = out & "Oct.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sqr(vD6): out = out & "Sqr.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Exp(vD6): out = out & "Exp.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Log(vD6): out = out & "Log.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Atn(vD6): out = out & "Atn.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sin(vD6): out = out & "Sin.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Cos(vD6): out = out & "Cos.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Tan(vD6): out = out & "Tan.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD6, -1): out = out & "Round.6.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD6, 0): out = out & "Round.6.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD6, 1): out = out & "Round.6.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD6, 2): out = out & "Round.6.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD6, "0.00"): out = out & "Format.6=" & Rz(Err.Number, r) & vbLf
  L007 = out
End Function
Function L008() As String
  Dim r As Variant, out As String, vA0 As Variant, vA1 As Variant, vA2 As Variant, vA3 As Variant, vA4 As Variant, vA5 As Variant, vA6 As Variant, vA7 As Variant, vA8 As Variant, vA9 As Variant, vA10 As Variant, vA11 As Variant, vD0 As Variant, vD1 As Variant, vD2 As Variant, vD3 As Variant, vD4 As Variant, vD5 As Variant, vD6 As Variant, vD7 As Variant, vD8 As Variant, vD9 As Variant, vD10 As Variant, vD11 As Variant, vD12 As Variant, vD13 As Variant, vD14 As Variant, vD15 As Variant, vD16 As Variant, vD17 As Variant
  On Error Resume Next
  vA0 = "": vA1 = "a": vA2 = "abc": vA3 = "Hello World": vA4 = "  pad  ": vA5 = "12": vA6 = "-3.5": vA7 = "1e3": vA8 = "&H1F": vA9 = "abc abc": vA10 = Null: vA11 = Empty
  vD0 = 0: vD1 = 1: vD2 = -1: vD3 = 0.5: vD4 = -0.5: vD5 = 1.5: vD6 = 2.5: vD7 = -1.5: vD8 = -2.5: vD9 = 3.5: vD10 = 100.5: vD11 = 255.5: vD12 = 32767.5: vD13 = 2147483647.5: vD14 = 1E+10: vD15 = 1E+20: vD16 = 123456.789: vD17 = -123456.789
  Err.Clear: r = Empty: r = Format(vD6, "General Number"): out = out & "FormatG.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD6, "#,##0.0"): out = out & "FormatP.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD6, "0.0E+00"): out = out & "FormatC.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD6, "0%"): out = out & "FormatPct.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(CStr(vD6)): out = out & "Val.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Int(vD7): out = out & "Int.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Fix(vD7): out = out & "Fix.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Abs(vD7): out = out & "Abs.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sgn(vD7): out = out & "Sgn.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vD7): out = out & "CInt.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vD7): out = out & "CLng.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vD7): out = out & "CByte.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vD7): out = out & "CDbl.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vD7): out = out & "CSng.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vD7): out = out & "CCur.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vD7): out = out & "CStr.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vD7): out = out & "CBool.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Hex(vD7): out = out & "Hex.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Oct(vD7): out = out & "Oct.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sqr(vD7): out = out & "Sqr.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Exp(vD7): out = out & "Exp.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Log(vD7): out = out & "Log.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Atn(vD7): out = out & "Atn.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sin(vD7): out = out & "Sin.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Cos(vD7): out = out & "Cos.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Tan(vD7): out = out & "Tan.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD7, -1): out = out & "Round.7.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD7, 0): out = out & "Round.7.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD7, 1): out = out & "Round.7.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD7, 2): out = out & "Round.7.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD7, "0.00"): out = out & "Format.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD7, "General Number"): out = out & "FormatG.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD7, "#,##0.0"): out = out & "FormatP.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD7, "0.0E+00"): out = out & "FormatC.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD7, "0%"): out = out & "FormatPct.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(CStr(vD7)): out = out & "Val.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Int(vD8): out = out & "Int.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Fix(vD8): out = out & "Fix.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Abs(vD8): out = out & "Abs.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sgn(vD8): out = out & "Sgn.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vD8): out = out & "CInt.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vD8): out = out & "CLng.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vD8): out = out & "CByte.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vD8): out = out & "CDbl.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vD8): out = out & "CSng.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vD8): out = out & "CCur.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vD8): out = out & "CStr.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vD8): out = out & "CBool.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Hex(vD8): out = out & "Hex.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Oct(vD8): out = out & "Oct.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sqr(vD8): out = out & "Sqr.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Exp(vD8): out = out & "Exp.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Log(vD8): out = out & "Log.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Atn(vD8): out = out & "Atn.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sin(vD8): out = out & "Sin.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Cos(vD8): out = out & "Cos.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Tan(vD8): out = out & "Tan.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD8, -1): out = out & "Round.8.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD8, 0): out = out & "Round.8.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD8, 1): out = out & "Round.8.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD8, 2): out = out & "Round.8.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD8, "0.00"): out = out & "Format.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD8, "General Number"): out = out & "FormatG.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD8, "#,##0.0"): out = out & "FormatP.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD8, "0.0E+00"): out = out & "FormatC.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD8, "0%"): out = out & "FormatPct.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(CStr(vD8)): out = out & "Val.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Int(vD9): out = out & "Int.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Fix(vD9): out = out & "Fix.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Abs(vD9): out = out & "Abs.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sgn(vD9): out = out & "Sgn.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vD9): out = out & "CInt.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vD9): out = out & "CLng.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vD9): out = out & "CByte.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vD9): out = out & "CDbl.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vD9): out = out & "CSng.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vD9): out = out & "CCur.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vD9): out = out & "CStr.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vD9): out = out & "CBool.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Hex(vD9): out = out & "Hex.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Oct(vD9): out = out & "Oct.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sqr(vD9): out = out & "Sqr.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Exp(vD9): out = out & "Exp.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Log(vD9): out = out & "Log.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Atn(vD9): out = out & "Atn.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sin(vD9): out = out & "Sin.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Cos(vD9): out = out & "Cos.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Tan(vD9): out = out & "Tan.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD9, -1): out = out & "Round.9.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD9, 0): out = out & "Round.9.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD9, 1): out = out & "Round.9.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD9, 2): out = out & "Round.9.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD9, "0.00"): out = out & "Format.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD9, "General Number"): out = out & "FormatG.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD9, "#,##0.0"): out = out & "FormatP.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD9, "0.0E+00"): out = out & "FormatC.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD9, "0%"): out = out & "FormatPct.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(CStr(vD9)): out = out & "Val.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Int(vD10): out = out & "Int.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Fix(vD10): out = out & "Fix.10=" & Rz(Err.Number, r) & vbLf
  L008 = out
End Function
Function L009() As String
  Dim r As Variant, out As String, vA0 As Variant, vA1 As Variant, vA2 As Variant, vA3 As Variant, vA4 As Variant, vA5 As Variant, vA6 As Variant, vA7 As Variant, vA8 As Variant, vA9 As Variant, vA10 As Variant, vA11 As Variant, vD0 As Variant, vD1 As Variant, vD2 As Variant, vD3 As Variant, vD4 As Variant, vD5 As Variant, vD6 As Variant, vD7 As Variant, vD8 As Variant, vD9 As Variant, vD10 As Variant, vD11 As Variant, vD12 As Variant, vD13 As Variant, vD14 As Variant, vD15 As Variant, vD16 As Variant, vD17 As Variant
  On Error Resume Next
  vA0 = "": vA1 = "a": vA2 = "abc": vA3 = "Hello World": vA4 = "  pad  ": vA5 = "12": vA6 = "-3.5": vA7 = "1e3": vA8 = "&H1F": vA9 = "abc abc": vA10 = Null: vA11 = Empty
  vD0 = 0: vD1 = 1: vD2 = -1: vD3 = 0.5: vD4 = -0.5: vD5 = 1.5: vD6 = 2.5: vD7 = -1.5: vD8 = -2.5: vD9 = 3.5: vD10 = 100.5: vD11 = 255.5: vD12 = 32767.5: vD13 = 2147483647.5: vD14 = 1E+10: vD15 = 1E+20: vD16 = 123456.789: vD17 = -123456.789
  Err.Clear: r = Empty: r = Abs(vD10): out = out & "Abs.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sgn(vD10): out = out & "Sgn.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vD10): out = out & "CInt.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vD10): out = out & "CLng.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vD10): out = out & "CByte.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vD10): out = out & "CDbl.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vD10): out = out & "CSng.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vD10): out = out & "CCur.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vD10): out = out & "CStr.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vD10): out = out & "CBool.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Hex(vD10): out = out & "Hex.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Oct(vD10): out = out & "Oct.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sqr(vD10): out = out & "Sqr.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Exp(vD10): out = out & "Exp.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Log(vD10): out = out & "Log.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Atn(vD10): out = out & "Atn.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sin(vD10): out = out & "Sin.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Cos(vD10): out = out & "Cos.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Tan(vD10): out = out & "Tan.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD10, -1): out = out & "Round.10.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD10, 0): out = out & "Round.10.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD10, 1): out = out & "Round.10.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD10, 2): out = out & "Round.10.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD10, "0.00"): out = out & "Format.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD10, "General Number"): out = out & "FormatG.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD10, "#,##0.0"): out = out & "FormatP.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD10, "0.0E+00"): out = out & "FormatC.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD10, "0%"): out = out & "FormatPct.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(CStr(vD10)): out = out & "Val.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Int(vD11): out = out & "Int.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Fix(vD11): out = out & "Fix.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Abs(vD11): out = out & "Abs.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sgn(vD11): out = out & "Sgn.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vD11): out = out & "CInt.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vD11): out = out & "CLng.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vD11): out = out & "CByte.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vD11): out = out & "CDbl.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vD11): out = out & "CSng.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vD11): out = out & "CCur.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vD11): out = out & "CStr.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vD11): out = out & "CBool.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Hex(vD11): out = out & "Hex.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Oct(vD11): out = out & "Oct.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sqr(vD11): out = out & "Sqr.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Exp(vD11): out = out & "Exp.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Log(vD11): out = out & "Log.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Atn(vD11): out = out & "Atn.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sin(vD11): out = out & "Sin.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Cos(vD11): out = out & "Cos.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Tan(vD11): out = out & "Tan.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD11, -1): out = out & "Round.11.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD11, 0): out = out & "Round.11.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD11, 1): out = out & "Round.11.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD11, 2): out = out & "Round.11.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD11, "0.00"): out = out & "Format.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD11, "General Number"): out = out & "FormatG.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD11, "#,##0.0"): out = out & "FormatP.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD11, "0.0E+00"): out = out & "FormatC.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD11, "0%"): out = out & "FormatPct.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(CStr(vD11)): out = out & "Val.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Int(vD12): out = out & "Int.12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Fix(vD12): out = out & "Fix.12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Abs(vD12): out = out & "Abs.12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sgn(vD12): out = out & "Sgn.12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vD12): out = out & "CInt.12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vD12): out = out & "CLng.12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vD12): out = out & "CByte.12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vD12): out = out & "CDbl.12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vD12): out = out & "CSng.12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vD12): out = out & "CCur.12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vD12): out = out & "CStr.12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vD12): out = out & "CBool.12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Hex(vD12): out = out & "Hex.12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Oct(vD12): out = out & "Oct.12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sqr(vD12): out = out & "Sqr.12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Exp(vD12): out = out & "Exp.12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Log(vD12): out = out & "Log.12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Atn(vD12): out = out & "Atn.12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sin(vD12): out = out & "Sin.12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Cos(vD12): out = out & "Cos.12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Tan(vD12): out = out & "Tan.12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD12, -1): out = out & "Round.12.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD12, 0): out = out & "Round.12.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD12, 1): out = out & "Round.12.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD12, 2): out = out & "Round.12.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD12, "0.00"): out = out & "Format.12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD12, "General Number"): out = out & "FormatG.12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD12, "#,##0.0"): out = out & "FormatP.12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD12, "0.0E+00"): out = out & "FormatC.12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD12, "0%"): out = out & "FormatPct.12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(CStr(vD12)): out = out & "Val.12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Int(vD13): out = out & "Int.13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Fix(vD13): out = out & "Fix.13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Abs(vD13): out = out & "Abs.13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sgn(vD13): out = out & "Sgn.13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vD13): out = out & "CInt.13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vD13): out = out & "CLng.13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vD13): out = out & "CByte.13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vD13): out = out & "CDbl.13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vD13): out = out & "CSng.13=" & Rz(Err.Number, r) & vbLf
  L009 = out
End Function
Function L010() As String
  Dim r As Variant, out As String, vA0 As Variant, vA1 As Variant, vA2 As Variant, vA3 As Variant, vA4 As Variant, vA5 As Variant, vA6 As Variant, vA7 As Variant, vA8 As Variant, vA9 As Variant, vA10 As Variant, vA11 As Variant, vD0 As Variant, vD1 As Variant, vD2 As Variant, vD3 As Variant, vD4 As Variant, vD5 As Variant, vD6 As Variant, vD7 As Variant, vD8 As Variant, vD9 As Variant, vD10 As Variant, vD11 As Variant, vD12 As Variant, vD13 As Variant, vD14 As Variant, vD15 As Variant, vD16 As Variant, vD17 As Variant
  On Error Resume Next
  vA0 = "": vA1 = "a": vA2 = "abc": vA3 = "Hello World": vA4 = "  pad  ": vA5 = "12": vA6 = "-3.5": vA7 = "1e3": vA8 = "&H1F": vA9 = "abc abc": vA10 = Null: vA11 = Empty
  vD0 = 0: vD1 = 1: vD2 = -1: vD3 = 0.5: vD4 = -0.5: vD5 = 1.5: vD6 = 2.5: vD7 = -1.5: vD8 = -2.5: vD9 = 3.5: vD10 = 100.5: vD11 = 255.5: vD12 = 32767.5: vD13 = 2147483647.5: vD14 = 1E+10: vD15 = 1E+20: vD16 = 123456.789: vD17 = -123456.789
  Err.Clear: r = Empty: r = CCur(vD13): out = out & "CCur.13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vD13): out = out & "CStr.13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vD13): out = out & "CBool.13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Hex(vD13): out = out & "Hex.13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Oct(vD13): out = out & "Oct.13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sqr(vD13): out = out & "Sqr.13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Exp(vD13): out = out & "Exp.13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Log(vD13): out = out & "Log.13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Atn(vD13): out = out & "Atn.13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sin(vD13): out = out & "Sin.13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Cos(vD13): out = out & "Cos.13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Tan(vD13): out = out & "Tan.13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD13, -1): out = out & "Round.13.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD13, 0): out = out & "Round.13.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD13, 1): out = out & "Round.13.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD13, 2): out = out & "Round.13.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD13, "0.00"): out = out & "Format.13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD13, "General Number"): out = out & "FormatG.13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD13, "#,##0.0"): out = out & "FormatP.13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD13, "0.0E+00"): out = out & "FormatC.13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD13, "0%"): out = out & "FormatPct.13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(CStr(vD13)): out = out & "Val.13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Int(vD14): out = out & "Int.14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Fix(vD14): out = out & "Fix.14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Abs(vD14): out = out & "Abs.14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sgn(vD14): out = out & "Sgn.14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vD14): out = out & "CInt.14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vD14): out = out & "CLng.14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vD14): out = out & "CByte.14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vD14): out = out & "CDbl.14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vD14): out = out & "CSng.14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vD14): out = out & "CCur.14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vD14): out = out & "CStr.14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vD14): out = out & "CBool.14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Hex(vD14): out = out & "Hex.14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Oct(vD14): out = out & "Oct.14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sqr(vD14): out = out & "Sqr.14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Exp(vD14): out = out & "Exp.14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Log(vD14): out = out & "Log.14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Atn(vD14): out = out & "Atn.14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sin(vD14): out = out & "Sin.14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Cos(vD14): out = out & "Cos.14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Tan(vD14): out = out & "Tan.14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD14, -1): out = out & "Round.14.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD14, 0): out = out & "Round.14.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD14, 1): out = out & "Round.14.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD14, 2): out = out & "Round.14.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD14, "0.00"): out = out & "Format.14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD14, "General Number"): out = out & "FormatG.14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD14, "#,##0.0"): out = out & "FormatP.14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD14, "0.0E+00"): out = out & "FormatC.14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD14, "0%"): out = out & "FormatPct.14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(CStr(vD14)): out = out & "Val.14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Int(vD15): out = out & "Int.15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Fix(vD15): out = out & "Fix.15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Abs(vD15): out = out & "Abs.15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sgn(vD15): out = out & "Sgn.15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vD15): out = out & "CInt.15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vD15): out = out & "CLng.15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vD15): out = out & "CByte.15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vD15): out = out & "CDbl.15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vD15): out = out & "CSng.15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vD15): out = out & "CCur.15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vD15): out = out & "CStr.15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vD15): out = out & "CBool.15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Hex(vD15): out = out & "Hex.15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Oct(vD15): out = out & "Oct.15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sqr(vD15): out = out & "Sqr.15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Exp(vD15): out = out & "Exp.15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Log(vD15): out = out & "Log.15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Atn(vD15): out = out & "Atn.15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sin(vD15): out = out & "Sin.15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Cos(vD15): out = out & "Cos.15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Tan(vD15): out = out & "Tan.15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD15, -1): out = out & "Round.15.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD15, 0): out = out & "Round.15.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD15, 1): out = out & "Round.15.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD15, 2): out = out & "Round.15.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD15, "0.00"): out = out & "Format.15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD15, "General Number"): out = out & "FormatG.15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD15, "#,##0.0"): out = out & "FormatP.15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD15, "0.0E+00"): out = out & "FormatC.15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD15, "0%"): out = out & "FormatPct.15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(CStr(vD15)): out = out & "Val.15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Int(vD16): out = out & "Int.16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Fix(vD16): out = out & "Fix.16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Abs(vD16): out = out & "Abs.16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sgn(vD16): out = out & "Sgn.16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vD16): out = out & "CInt.16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vD16): out = out & "CLng.16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vD16): out = out & "CByte.16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vD16): out = out & "CDbl.16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vD16): out = out & "CSng.16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vD16): out = out & "CCur.16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vD16): out = out & "CStr.16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vD16): out = out & "CBool.16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Hex(vD16): out = out & "Hex.16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Oct(vD16): out = out & "Oct.16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sqr(vD16): out = out & "Sqr.16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Exp(vD16): out = out & "Exp.16=" & Rz(Err.Number, r) & vbLf
  L010 = out
End Function
Function L011() As String
  Dim r As Variant, out As String, vA0 As Variant, vA1 As Variant, vA2 As Variant, vA3 As Variant, vA4 As Variant, vA5 As Variant, vA6 As Variant, vA7 As Variant, vA8 As Variant, vA9 As Variant, vA10 As Variant, vA11 As Variant, vD0 As Variant, vD1 As Variant, vD2 As Variant, vD3 As Variant, vD4 As Variant, vD5 As Variant, vD6 As Variant, vD7 As Variant, vD8 As Variant, vD9 As Variant, vD10 As Variant, vD11 As Variant, vD12 As Variant, vD13 As Variant, vD14 As Variant, vD15 As Variant, vD16 As Variant, vD17 As Variant
  On Error Resume Next
  vA0 = "": vA1 = "a": vA2 = "abc": vA3 = "Hello World": vA4 = "  pad  ": vA5 = "12": vA6 = "-3.5": vA7 = "1e3": vA8 = "&H1F": vA9 = "abc abc": vA10 = Null: vA11 = Empty
  vD0 = 0: vD1 = 1: vD2 = -1: vD3 = 0.5: vD4 = -0.5: vD5 = 1.5: vD6 = 2.5: vD7 = -1.5: vD8 = -2.5: vD9 = 3.5: vD10 = 100.5: vD11 = 255.5: vD12 = 32767.5: vD13 = 2147483647.5: vD14 = 1E+10: vD15 = 1E+20: vD16 = 123456.789: vD17 = -123456.789
  Err.Clear: r = Empty: r = Log(vD16): out = out & "Log.16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Atn(vD16): out = out & "Atn.16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sin(vD16): out = out & "Sin.16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Cos(vD16): out = out & "Cos.16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Tan(vD16): out = out & "Tan.16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD16, -1): out = out & "Round.16.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD16, 0): out = out & "Round.16.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD16, 1): out = out & "Round.16.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD16, 2): out = out & "Round.16.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD16, "0.00"): out = out & "Format.16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD16, "General Number"): out = out & "FormatG.16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD16, "#,##0.0"): out = out & "FormatP.16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD16, "0.0E+00"): out = out & "FormatC.16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD16, "0%"): out = out & "FormatPct.16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(CStr(vD16)): out = out & "Val.16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Int(vD17): out = out & "Int.17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Fix(vD17): out = out & "Fix.17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Abs(vD17): out = out & "Abs.17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sgn(vD17): out = out & "Sgn.17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vD17): out = out & "CInt.17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vD17): out = out & "CLng.17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vD17): out = out & "CByte.17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vD17): out = out & "CDbl.17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vD17): out = out & "CSng.17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vD17): out = out & "CCur.17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vD17): out = out & "CStr.17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vD17): out = out & "CBool.17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Hex(vD17): out = out & "Hex.17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Oct(vD17): out = out & "Oct.17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sqr(vD17): out = out & "Sqr.17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Exp(vD17): out = out & "Exp.17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Log(vD17): out = out & "Log.17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Atn(vD17): out = out & "Atn.17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Sin(vD17): out = out & "Sin.17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Cos(vD17): out = out & "Cos.17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Tan(vD17): out = out & "Tan.17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD17, -1): out = out & "Round.17.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD17, 0): out = out & "Round.17.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD17, 1): out = out & "Round.17.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Round(vD17, 2): out = out & "Round.17.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD17, "0.00"): out = out & "Format.17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD17, "General Number"): out = out & "FormatG.17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD17, "#,##0.0"): out = out & "FormatP.17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD17, "0.0E+00"): out = out & "FormatC.17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Format(vD17, "0%"): out = out & "FormatPct.17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(CStr(vD17)): out = out & "Val.17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vA0): out = out & "CIntS.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vA0): out = out & "CLngS.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vA0): out = out & "CDblS.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vA0): out = out & "CBoolS.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vA0): out = out & "CByteS.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vA0): out = out & "CCurS.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vA0): out = out & "CSngS.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vA0): out = out & "CStrS.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(vA0): out = out & "ValS.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsNumeric(vA0): out = out & "IsNumericS.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsDate(vA0): out = out & "IsDateS.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsEmpty(vA0): out = out & "IsEmptyS.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsNull(vA0): out = out & "IsNullS.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vA1): out = out & "CIntS.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vA1): out = out & "CLngS.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vA1): out = out & "CDblS.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vA1): out = out & "CBoolS.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vA1): out = out & "CByteS.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vA1): out = out & "CCurS.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vA1): out = out & "CSngS.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vA1): out = out & "CStrS.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(vA1): out = out & "ValS.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsNumeric(vA1): out = out & "IsNumericS.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsDate(vA1): out = out & "IsDateS.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsEmpty(vA1): out = out & "IsEmptyS.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsNull(vA1): out = out & "IsNullS.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vA2): out = out & "CIntS.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vA2): out = out & "CLngS.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vA2): out = out & "CDblS.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vA2): out = out & "CBoolS.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vA2): out = out & "CByteS.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vA2): out = out & "CCurS.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vA2): out = out & "CSngS.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vA2): out = out & "CStrS.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(vA2): out = out & "ValS.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsNumeric(vA2): out = out & "IsNumericS.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsDate(vA2): out = out & "IsDateS.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsEmpty(vA2): out = out & "IsEmptyS.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsNull(vA2): out = out & "IsNullS.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vA3): out = out & "CIntS.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vA3): out = out & "CLngS.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vA3): out = out & "CDblS.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vA3): out = out & "CBoolS.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vA3): out = out & "CByteS.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vA3): out = out & "CCurS.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vA3): out = out & "CSngS.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vA3): out = out & "CStrS.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(vA3): out = out & "ValS.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsNumeric(vA3): out = out & "IsNumericS.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsDate(vA3): out = out & "IsDateS.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsEmpty(vA3): out = out & "IsEmptyS.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsNull(vA3): out = out & "IsNullS.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vA4): out = out & "CIntS.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vA4): out = out & "CLngS.4=" & Rz(Err.Number, r) & vbLf
  L011 = out
End Function
Function L012() As String
  Dim r As Variant, out As String, vA0 As Variant, vA1 As Variant, vA2 As Variant, vA3 As Variant, vA4 As Variant, vA5 As Variant, vA6 As Variant, vA7 As Variant, vA8 As Variant, vA9 As Variant, vA10 As Variant, vA11 As Variant, vD0 As Variant, vD1 As Variant, vD2 As Variant, vD3 As Variant, vD4 As Variant, vD5 As Variant, vD6 As Variant, vD7 As Variant, vD8 As Variant, vD9 As Variant, vD10 As Variant, vD11 As Variant, vD12 As Variant, vD13 As Variant, vD14 As Variant, vD15 As Variant, vD16 As Variant, vD17 As Variant
  On Error Resume Next
  vA0 = "": vA1 = "a": vA2 = "abc": vA3 = "Hello World": vA4 = "  pad  ": vA5 = "12": vA6 = "-3.5": vA7 = "1e3": vA8 = "&H1F": vA9 = "abc abc": vA10 = Null: vA11 = Empty
  vD0 = 0: vD1 = 1: vD2 = -1: vD3 = 0.5: vD4 = -0.5: vD5 = 1.5: vD6 = 2.5: vD7 = -1.5: vD8 = -2.5: vD9 = 3.5: vD10 = 100.5: vD11 = 255.5: vD12 = 32767.5: vD13 = 2147483647.5: vD14 = 1E+10: vD15 = 1E+20: vD16 = 123456.789: vD17 = -123456.789
  Err.Clear: r = Empty: r = CDbl(vA4): out = out & "CDblS.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vA4): out = out & "CBoolS.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vA4): out = out & "CByteS.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vA4): out = out & "CCurS.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vA4): out = out & "CSngS.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vA4): out = out & "CStrS.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(vA4): out = out & "ValS.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsNumeric(vA4): out = out & "IsNumericS.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsDate(vA4): out = out & "IsDateS.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsEmpty(vA4): out = out & "IsEmptyS.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsNull(vA4): out = out & "IsNullS.4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vA5): out = out & "CIntS.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vA5): out = out & "CLngS.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vA5): out = out & "CDblS.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vA5): out = out & "CBoolS.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vA5): out = out & "CByteS.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vA5): out = out & "CCurS.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vA5): out = out & "CSngS.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vA5): out = out & "CStrS.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(vA5): out = out & "ValS.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsNumeric(vA5): out = out & "IsNumericS.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsDate(vA5): out = out & "IsDateS.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsEmpty(vA5): out = out & "IsEmptyS.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsNull(vA5): out = out & "IsNullS.5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vA6): out = out & "CIntS.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vA6): out = out & "CLngS.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vA6): out = out & "CDblS.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vA6): out = out & "CBoolS.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vA6): out = out & "CByteS.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vA6): out = out & "CCurS.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vA6): out = out & "CSngS.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vA6): out = out & "CStrS.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(vA6): out = out & "ValS.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsNumeric(vA6): out = out & "IsNumericS.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsDate(vA6): out = out & "IsDateS.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsEmpty(vA6): out = out & "IsEmptyS.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsNull(vA6): out = out & "IsNullS.6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vA7): out = out & "CIntS.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vA7): out = out & "CLngS.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vA7): out = out & "CDblS.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vA7): out = out & "CBoolS.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vA7): out = out & "CByteS.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vA7): out = out & "CCurS.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vA7): out = out & "CSngS.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vA7): out = out & "CStrS.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(vA7): out = out & "ValS.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsNumeric(vA7): out = out & "IsNumericS.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsDate(vA7): out = out & "IsDateS.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsEmpty(vA7): out = out & "IsEmptyS.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsNull(vA7): out = out & "IsNullS.7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vA8): out = out & "CIntS.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vA8): out = out & "CLngS.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vA8): out = out & "CDblS.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vA8): out = out & "CBoolS.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vA8): out = out & "CByteS.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vA8): out = out & "CCurS.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vA8): out = out & "CSngS.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vA8): out = out & "CStrS.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(vA8): out = out & "ValS.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsNumeric(vA8): out = out & "IsNumericS.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsDate(vA8): out = out & "IsDateS.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsEmpty(vA8): out = out & "IsEmptyS.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsNull(vA8): out = out & "IsNullS.8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vA9): out = out & "CIntS.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vA9): out = out & "CLngS.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vA9): out = out & "CDblS.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vA9): out = out & "CBoolS.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vA9): out = out & "CByteS.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vA9): out = out & "CCurS.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vA9): out = out & "CSngS.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vA9): out = out & "CStrS.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(vA9): out = out & "ValS.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsNumeric(vA9): out = out & "IsNumericS.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsDate(vA9): out = out & "IsDateS.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsEmpty(vA9): out = out & "IsEmptyS.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsNull(vA9): out = out & "IsNullS.9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vA10): out = out & "CIntS.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vA10): out = out & "CLngS.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vA10): out = out & "CDblS.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vA10): out = out & "CBoolS.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vA10): out = out & "CByteS.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vA10): out = out & "CCurS.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vA10): out = out & "CSngS.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vA10): out = out & "CStrS.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(vA10): out = out & "ValS.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsNumeric(vA10): out = out & "IsNumericS.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsDate(vA10): out = out & "IsDateS.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsEmpty(vA10): out = out & "IsEmptyS.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsNull(vA10): out = out & "IsNullS.10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CInt(vA11): out = out & "CIntS.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CLng(vA11): out = out & "CLngS.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(vA11): out = out & "CDblS.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CBool(vA11): out = out & "CBoolS.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CByte(vA11): out = out & "CByteS.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CCur(vA11): out = out & "CCurS.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CSng(vA11): out = out & "CSngS.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CStr(vA11): out = out & "CStrS.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Val(vA11): out = out & "ValS.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsNumeric(vA11): out = out & "IsNumericS.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsDate(vA11): out = out & "IsDateS.11=" & Rz(Err.Number, r) & vbLf
  L012 = out
End Function
Function L013() As String
  Dim r As Variant, out As String, vA0 As Variant, vA1 As Variant, vA2 As Variant, vA3 As Variant, vA4 As Variant, vA5 As Variant, vA6 As Variant, vA7 As Variant, vA8 As Variant, vA9 As Variant, vA10 As Variant, vA11 As Variant, vD0 As Variant, vD1 As Variant, vD2 As Variant, vD3 As Variant, vD4 As Variant, vD5 As Variant, vD6 As Variant, vD7 As Variant, vD8 As Variant, vD9 As Variant, vD10 As Variant, vD11 As Variant, vD12 As Variant, vD13 As Variant, vD14 As Variant, vD15 As Variant, vD16 As Variant, vD17 As Variant
  On Error Resume Next
  vA0 = "": vA1 = "a": vA2 = "abc": vA3 = "Hello World": vA4 = "  pad  ": vA5 = "12": vA6 = "-3.5": vA7 = "1e3": vA8 = "&H1F": vA9 = "abc abc": vA10 = Null: vA11 = Empty
  vD0 = 0: vD1 = 1: vD2 = -1: vD3 = 0.5: vD4 = -0.5: vD5 = 1.5: vD6 = 2.5: vD7 = -1.5: vD8 = -2.5: vD9 = 3.5: vD10 = 100.5: vD11 = 255.5: vD12 = 32767.5: vD13 = 2147483647.5: vD14 = 1E+10: vD15 = 1E+20: vD16 = 123456.789: vD17 = -123456.789
  Err.Clear: r = Empty: r = IsEmpty(vA11): out = out & "IsEmptyS.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = IsNull(vA11): out = out & "IsNullS.11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 1 And 1: out = out & "And.1.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 1 Or 1: out = out & "Or.1.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 1 Xor 1: out = out & "Xor.1.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 1 Imp 1: out = out & "Imp.1.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 1 Eqv 1: out = out & "Eqv.1.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 1 And 2: out = out & "And.1.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 1 Or 2: out = out & "Or.1.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 1 Xor 2: out = out & "Xor.1.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 1 Imp 2: out = out & "Imp.1.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 1 Eqv 2: out = out & "Eqv.1.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 1 And 0: out = out & "And.1.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 1 Or 0: out = out & "Or.1.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 1 Xor 0: out = out & "Xor.1.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 1 Imp 0: out = out & "Imp.1.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 1 Eqv 0: out = out & "Eqv.1.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 1 And -1: out = out & "And.1.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 1 Or -1: out = out & "Or.1.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 1 Xor -1: out = out & "Xor.1.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 1 Imp -1: out = out & "Imp.1.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 1 Eqv -1: out = out & "Eqv.1.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 1 And 3: out = out & "And.1.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 1 Or 3: out = out & "Or.1.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 1 Xor 3: out = out & "Xor.1.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 1 Imp 3: out = out & "Imp.1.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 1 Eqv 3: out = out & "Eqv.1.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 2 And 1: out = out & "And.2.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 2 Or 1: out = out & "Or.2.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 2 Xor 1: out = out & "Xor.2.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 2 Imp 1: out = out & "Imp.2.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 2 Eqv 1: out = out & "Eqv.2.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 2 And 2: out = out & "And.2.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 2 Or 2: out = out & "Or.2.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 2 Xor 2: out = out & "Xor.2.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 2 Imp 2: out = out & "Imp.2.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 2 Eqv 2: out = out & "Eqv.2.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 2 And 0: out = out & "And.2.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 2 Or 0: out = out & "Or.2.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 2 Xor 0: out = out & "Xor.2.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 2 Imp 0: out = out & "Imp.2.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 2 Eqv 0: out = out & "Eqv.2.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 2 And -1: out = out & "And.2.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 2 Or -1: out = out & "Or.2.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 2 Xor -1: out = out & "Xor.2.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 2 Imp -1: out = out & "Imp.2.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 2 Eqv -1: out = out & "Eqv.2.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 2 And 3: out = out & "And.2.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 2 Or 3: out = out & "Or.2.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 2 Xor 3: out = out & "Xor.2.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 2 Imp 3: out = out & "Imp.2.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 2 Eqv 3: out = out & "Eqv.2.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 0 And 1: out = out & "And.0.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 0 Or 1: out = out & "Or.0.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 0 Xor 1: out = out & "Xor.0.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 0 Imp 1: out = out & "Imp.0.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 0 Eqv 1: out = out & "Eqv.0.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 0 And 2: out = out & "And.0.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 0 Or 2: out = out & "Or.0.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 0 Xor 2: out = out & "Xor.0.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 0 Imp 2: out = out & "Imp.0.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 0 Eqv 2: out = out & "Eqv.0.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 0 And 0: out = out & "And.0.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 0 Or 0: out = out & "Or.0.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 0 Xor 0: out = out & "Xor.0.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 0 Imp 0: out = out & "Imp.0.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 0 Eqv 0: out = out & "Eqv.0.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 0 And -1: out = out & "And.0.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 0 Or -1: out = out & "Or.0.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 0 Xor -1: out = out & "Xor.0.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 0 Imp -1: out = out & "Imp.0.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 0 Eqv -1: out = out & "Eqv.0.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 0 And 3: out = out & "And.0.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 0 Or 3: out = out & "Or.0.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 0 Xor 3: out = out & "Xor.0.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 0 Imp 3: out = out & "Imp.0.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 0 Eqv 3: out = out & "Eqv.0.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = -1 And 1: out = out & "And.-1.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = -1 Or 1: out = out & "Or.-1.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = -1 Xor 1: out = out & "Xor.-1.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = -1 Imp 1: out = out & "Imp.-1.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = -1 Eqv 1: out = out & "Eqv.-1.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = -1 And 2: out = out & "And.-1.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = -1 Or 2: out = out & "Or.-1.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = -1 Xor 2: out = out & "Xor.-1.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = -1 Imp 2: out = out & "Imp.-1.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = -1 Eqv 2: out = out & "Eqv.-1.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = -1 And 0: out = out & "And.-1.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = -1 Or 0: out = out & "Or.-1.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = -1 Xor 0: out = out & "Xor.-1.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = -1 Imp 0: out = out & "Imp.-1.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = -1 Eqv 0: out = out & "Eqv.-1.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = -1 And -1: out = out & "And.-1.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = -1 Or -1: out = out & "Or.-1.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = -1 Xor -1: out = out & "Xor.-1.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = -1 Imp -1: out = out & "Imp.-1.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = -1 Eqv -1: out = out & "Eqv.-1.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = -1 And 3: out = out & "And.-1.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = -1 Or 3: out = out & "Or.-1.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = -1 Xor 3: out = out & "Xor.-1.3=" & Rz(Err.Number, r) & vbLf
  L013 = out
End Function
Function L014() As String
  Dim r As Variant, out As String, vA0 As Variant, vA1 As Variant, vA2 As Variant, vA3 As Variant, vA4 As Variant, vA5 As Variant, vA6 As Variant, vA7 As Variant, vA8 As Variant, vA9 As Variant, vA10 As Variant, vA11 As Variant, vD0 As Variant, vD1 As Variant, vD2 As Variant, vD3 As Variant, vD4 As Variant, vD5 As Variant, vD6 As Variant, vD7 As Variant, vD8 As Variant, vD9 As Variant, vD10 As Variant, vD11 As Variant, vD12 As Variant, vD13 As Variant, vD14 As Variant, vD15 As Variant, vD16 As Variant, vD17 As Variant
  On Error Resume Next
  vA0 = "": vA1 = "a": vA2 = "abc": vA3 = "Hello World": vA4 = "  pad  ": vA5 = "12": vA6 = "-3.5": vA7 = "1e3": vA8 = "&H1F": vA9 = "abc abc": vA10 = Null: vA11 = Empty
  vD0 = 0: vD1 = 1: vD2 = -1: vD3 = 0.5: vD4 = -0.5: vD5 = 1.5: vD6 = 2.5: vD7 = -1.5: vD8 = -2.5: vD9 = 3.5: vD10 = 100.5: vD11 = 255.5: vD12 = 32767.5: vD13 = 2147483647.5: vD14 = 1E+10: vD15 = 1E+20: vD16 = 123456.789: vD17 = -123456.789
  Err.Clear: r = Empty: r = -1 Imp 3: out = out & "Imp.-1.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = -1 Eqv 3: out = out & "Eqv.-1.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 255 And 1: out = out & "And.255.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 255 Or 1: out = out & "Or.255.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 255 Xor 1: out = out & "Xor.255.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 255 Imp 1: out = out & "Imp.255.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 255 Eqv 1: out = out & "Eqv.255.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 255 And 2: out = out & "And.255.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 255 Or 2: out = out & "Or.255.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 255 Xor 2: out = out & "Xor.255.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 255 Imp 2: out = out & "Imp.255.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 255 Eqv 2: out = out & "Eqv.255.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 255 And 0: out = out & "And.255.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 255 Or 0: out = out & "Or.255.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 255 Xor 0: out = out & "Xor.255.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 255 Imp 0: out = out & "Imp.255.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 255 Eqv 0: out = out & "Eqv.255.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 255 And -1: out = out & "And.255.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 255 Or -1: out = out & "Or.255.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 255 Xor -1: out = out & "Xor.255.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 255 Imp -1: out = out & "Imp.255.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 255 Eqv -1: out = out & "Eqv.255.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 255 And 3: out = out & "And.255.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 255 Or 3: out = out & "Or.255.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 255 Xor 3: out = out & "Xor.255.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 255 Imp 3: out = out & "Imp.255.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 255 Eqv 3: out = out & "Eqv.255.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 256 And 1: out = out & "And.256.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 256 Or 1: out = out & "Or.256.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 256 Xor 1: out = out & "Xor.256.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 256 Imp 1: out = out & "Imp.256.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 256 Eqv 1: out = out & "Eqv.256.1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 256 And 2: out = out & "And.256.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 256 Or 2: out = out & "Or.256.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 256 Xor 2: out = out & "Xor.256.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 256 Imp 2: out = out & "Imp.256.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 256 Eqv 2: out = out & "Eqv.256.2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 256 And 0: out = out & "And.256.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 256 Or 0: out = out & "Or.256.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 256 Xor 0: out = out & "Xor.256.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 256 Imp 0: out = out & "Imp.256.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 256 Eqv 0: out = out & "Eqv.256.0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 256 And -1: out = out & "And.256.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 256 Or -1: out = out & "Or.256.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 256 Xor -1: out = out & "Xor.256.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 256 Imp -1: out = out & "Imp.256.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 256 Eqv -1: out = out & "Eqv.256.-1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 256 And 3: out = out & "And.256.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 256 Or 3: out = out & "Or.256.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 256 Xor 3: out = out & "Xor.256.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 256 Imp 3: out = out & "Imp.256.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = 256 Eqv 3: out = out & "Eqv.256.3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = True And True: out = out & "AndB.True.True=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = True Or True: out = out & "OrB.True.True=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = True And False: out = out & "AndB.True.False=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = True Or False: out = out & "OrB.True.False=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = True And Null: out = out & "AndB.True.Null=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = True Or Null: out = out & "OrB.True.Null=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = False And True: out = out & "AndB.False.True=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = False Or True: out = out & "OrB.False.True=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = False And False: out = out & "AndB.False.False=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = False Or False: out = out & "OrB.False.False=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = False And Null: out = out & "AndB.False.Null=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = False Or Null: out = out & "OrB.False.Null=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Null And True: out = out & "AndB.Null.True=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Null Or True: out = out & "OrB.Null.True=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Null And False: out = out & "AndB.Null.False=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Null Or False: out = out & "OrB.Null.False=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Null And Null: out = out & "AndB.Null.Null=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Null Or Null: out = out & "OrB.Null.Null=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Year(#1/1/2000#): out = out & "Year.1_1_2000=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Month(#1/1/2000#): out = out & "Month.1_1_2000=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Day(#1/1/2000#): out = out & "Day.1_1_2000=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Hour(#1/1/2000#): out = out & "Hour.1_1_2000=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Minute(#1/1/2000#): out = out & "Minute.1_1_2000=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Second(#1/1/2000#): out = out & "Second.1_1_2000=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Weekday(#1/1/2000#): out = out & "Weekday.1_1_2000=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(#1/1/2000#): out = out & "CDbl.1_1_2000=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("d", 13, #1/1/2000#)): out = out & "DateAdd.1_1_2000.d=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("d", #1/1/2000#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.1_1_2000.d=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("d", #1/1/2000#): out = out & "DatePart.1_1_2000.d=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("m", 13, #1/1/2000#)): out = out & "DateAdd.1_1_2000.m=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("m", #1/1/2000#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.1_1_2000.m=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("m", #1/1/2000#): out = out & "DatePart.1_1_2000.m=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("yyyy", 13, #1/1/2000#)): out = out & "DateAdd.1_1_2000.yyyy=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("yyyy", #1/1/2000#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.1_1_2000.yyyy=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("yyyy", #1/1/2000#): out = out & "DatePart.1_1_2000.yyyy=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("h", 13, #1/1/2000#)): out = out & "DateAdd.1_1_2000.h=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("h", #1/1/2000#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.1_1_2000.h=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("h", #1/1/2000#): out = out & "DatePart.1_1_2000.h=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("n", 13, #1/1/2000#)): out = out & "DateAdd.1_1_2000.n=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("n", #1/1/2000#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.1_1_2000.n=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("n", #1/1/2000#): out = out & "DatePart.1_1_2000.n=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("s", 13, #1/1/2000#)): out = out & "DateAdd.1_1_2000.s=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("s", #1/1/2000#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.1_1_2000.s=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("s", #1/1/2000#): out = out & "DatePart.1_1_2000.s=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("w", 13, #1/1/2000#)): out = out & "DateAdd.1_1_2000.w=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("w", #1/1/2000#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.1_1_2000.w=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("w", #1/1/2000#): out = out & "DatePart.1_1_2000.w=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("ww", 13, #1/1/2000#)): out = out & "DateAdd.1_1_2000.ww=" & Rz(Err.Number, r) & vbLf
  L014 = out
End Function
Function L015() As String
  Dim r As Variant, out As String, vA0 As Variant, vA1 As Variant, vA2 As Variant, vA3 As Variant, vA4 As Variant, vA5 As Variant, vA6 As Variant, vA7 As Variant, vA8 As Variant, vA9 As Variant, vA10 As Variant, vA11 As Variant, vD0 As Variant, vD1 As Variant, vD2 As Variant, vD3 As Variant, vD4 As Variant, vD5 As Variant, vD6 As Variant, vD7 As Variant, vD8 As Variant, vD9 As Variant, vD10 As Variant, vD11 As Variant, vD12 As Variant, vD13 As Variant, vD14 As Variant, vD15 As Variant, vD16 As Variant, vD17 As Variant
  On Error Resume Next
  vA0 = "": vA1 = "a": vA2 = "abc": vA3 = "Hello World": vA4 = "  pad  ": vA5 = "12": vA6 = "-3.5": vA7 = "1e3": vA8 = "&H1F": vA9 = "abc abc": vA10 = Null: vA11 = Empty
  vD0 = 0: vD1 = 1: vD2 = -1: vD3 = 0.5: vD4 = -0.5: vD5 = 1.5: vD6 = 2.5: vD7 = -1.5: vD8 = -2.5: vD9 = 3.5: vD10 = 100.5: vD11 = 255.5: vD12 = 32767.5: vD13 = 2147483647.5: vD14 = 1E+10: vD15 = 1E+20: vD16 = 123456.789: vD17 = -123456.789
  Err.Clear: r = Empty: r = DateDiff("ww", #1/1/2000#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.1_1_2000.ww=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("ww", #1/1/2000#): out = out & "DatePart.1_1_2000.ww=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("q", 13, #1/1/2000#)): out = out & "DateAdd.1_1_2000.q=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("q", #1/1/2000#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.1_1_2000.q=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("q", #1/1/2000#): out = out & "DatePart.1_1_2000.q=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Year(#12/31/1999 11:59:59 PM#): out = out & "Year.12_31_1999115959PM=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Month(#12/31/1999 11:59:59 PM#): out = out & "Month.12_31_1999115959PM=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Day(#12/31/1999 11:59:59 PM#): out = out & "Day.12_31_1999115959PM=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Hour(#12/31/1999 11:59:59 PM#): out = out & "Hour.12_31_1999115959PM=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Minute(#12/31/1999 11:59:59 PM#): out = out & "Minute.12_31_1999115959PM=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Second(#12/31/1999 11:59:59 PM#): out = out & "Second.12_31_1999115959PM=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Weekday(#12/31/1999 11:59:59 PM#): out = out & "Weekday.12_31_1999115959PM=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(#12/31/1999 11:59:59 PM#): out = out & "CDbl.12_31_1999115959PM=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("d", 13, #12/31/1999 11:59:59 PM#)): out = out & "DateAdd.12_31_1999115959PM.d=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("d", #12/31/1999 11:59:59 PM#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.12_31_1999115959PM.d=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("d", #12/31/1999 11:59:59 PM#): out = out & "DatePart.12_31_1999115959PM.d=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("m", 13, #12/31/1999 11:59:59 PM#)): out = out & "DateAdd.12_31_1999115959PM.m=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("m", #12/31/1999 11:59:59 PM#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.12_31_1999115959PM.m=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("m", #12/31/1999 11:59:59 PM#): out = out & "DatePart.12_31_1999115959PM.m=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("yyyy", 13, #12/31/1999 11:59:59 PM#)): out = out & "DateAdd.12_31_1999115959PM.yyyy=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("yyyy", #12/31/1999 11:59:59 PM#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.12_31_1999115959PM.yyyy=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("yyyy", #12/31/1999 11:59:59 PM#): out = out & "DatePart.12_31_1999115959PM.yyyy=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("h", 13, #12/31/1999 11:59:59 PM#)): out = out & "DateAdd.12_31_1999115959PM.h=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("h", #12/31/1999 11:59:59 PM#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.12_31_1999115959PM.h=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("h", #12/31/1999 11:59:59 PM#): out = out & "DatePart.12_31_1999115959PM.h=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("n", 13, #12/31/1999 11:59:59 PM#)): out = out & "DateAdd.12_31_1999115959PM.n=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("n", #12/31/1999 11:59:59 PM#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.12_31_1999115959PM.n=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("n", #12/31/1999 11:59:59 PM#): out = out & "DatePart.12_31_1999115959PM.n=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("s", 13, #12/31/1999 11:59:59 PM#)): out = out & "DateAdd.12_31_1999115959PM.s=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("s", #12/31/1999 11:59:59 PM#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.12_31_1999115959PM.s=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("s", #12/31/1999 11:59:59 PM#): out = out & "DatePart.12_31_1999115959PM.s=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("w", 13, #12/31/1999 11:59:59 PM#)): out = out & "DateAdd.12_31_1999115959PM.w=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("w", #12/31/1999 11:59:59 PM#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.12_31_1999115959PM.w=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("w", #12/31/1999 11:59:59 PM#): out = out & "DatePart.12_31_1999115959PM.w=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("ww", 13, #12/31/1999 11:59:59 PM#)): out = out & "DateAdd.12_31_1999115959PM.ww=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("ww", #12/31/1999 11:59:59 PM#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.12_31_1999115959PM.ww=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("ww", #12/31/1999 11:59:59 PM#): out = out & "DatePart.12_31_1999115959PM.ww=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("q", 13, #12/31/1999 11:59:59 PM#)): out = out & "DateAdd.12_31_1999115959PM.q=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("q", #12/31/1999 11:59:59 PM#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.12_31_1999115959PM.q=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("q", #12/31/1999 11:59:59 PM#): out = out & "DatePart.12_31_1999115959PM.q=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Year(#2/29/2000#): out = out & "Year.2_29_2000=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Month(#2/29/2000#): out = out & "Month.2_29_2000=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Day(#2/29/2000#): out = out & "Day.2_29_2000=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Hour(#2/29/2000#): out = out & "Hour.2_29_2000=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Minute(#2/29/2000#): out = out & "Minute.2_29_2000=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Second(#2/29/2000#): out = out & "Second.2_29_2000=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Weekday(#2/29/2000#): out = out & "Weekday.2_29_2000=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(#2/29/2000#): out = out & "CDbl.2_29_2000=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("d", 13, #2/29/2000#)): out = out & "DateAdd.2_29_2000.d=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("d", #2/29/2000#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.2_29_2000.d=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("d", #2/29/2000#): out = out & "DatePart.2_29_2000.d=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("m", 13, #2/29/2000#)): out = out & "DateAdd.2_29_2000.m=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("m", #2/29/2000#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.2_29_2000.m=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("m", #2/29/2000#): out = out & "DatePart.2_29_2000.m=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("yyyy", 13, #2/29/2000#)): out = out & "DateAdd.2_29_2000.yyyy=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("yyyy", #2/29/2000#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.2_29_2000.yyyy=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("yyyy", #2/29/2000#): out = out & "DatePart.2_29_2000.yyyy=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("h", 13, #2/29/2000#)): out = out & "DateAdd.2_29_2000.h=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("h", #2/29/2000#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.2_29_2000.h=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("h", #2/29/2000#): out = out & "DatePart.2_29_2000.h=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("n", 13, #2/29/2000#)): out = out & "DateAdd.2_29_2000.n=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("n", #2/29/2000#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.2_29_2000.n=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("n", #2/29/2000#): out = out & "DatePart.2_29_2000.n=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("s", 13, #2/29/2000#)): out = out & "DateAdd.2_29_2000.s=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("s", #2/29/2000#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.2_29_2000.s=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("s", #2/29/2000#): out = out & "DatePart.2_29_2000.s=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("w", 13, #2/29/2000#)): out = out & "DateAdd.2_29_2000.w=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("w", #2/29/2000#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.2_29_2000.w=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("w", #2/29/2000#): out = out & "DatePart.2_29_2000.w=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("ww", 13, #2/29/2000#)): out = out & "DateAdd.2_29_2000.ww=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("ww", #2/29/2000#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.2_29_2000.ww=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("ww", #2/29/2000#): out = out & "DatePart.2_29_2000.ww=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("q", 13, #2/29/2000#)): out = out & "DateAdd.2_29_2000.q=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("q", #2/29/2000#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.2_29_2000.q=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("q", #2/29/2000#): out = out & "DatePart.2_29_2000.q=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Year(#3/1/1900#): out = out & "Year.3_1_1900=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Month(#3/1/1900#): out = out & "Month.3_1_1900=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Day(#3/1/1900#): out = out & "Day.3_1_1900=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Hour(#3/1/1900#): out = out & "Hour.3_1_1900=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Minute(#3/1/1900#): out = out & "Minute.3_1_1900=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Second(#3/1/1900#): out = out & "Second.3_1_1900=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = Weekday(#3/1/1900#): out = out & "Weekday.3_1_1900=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(#3/1/1900#): out = out & "CDbl.3_1_1900=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("d", 13, #3/1/1900#)): out = out & "DateAdd.3_1_1900.d=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("d", #3/1/1900#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.3_1_1900.d=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("d", #3/1/1900#): out = out & "DatePart.3_1_1900.d=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("m", 13, #3/1/1900#)): out = out & "DateAdd.3_1_1900.m=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("m", #3/1/1900#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.3_1_1900.m=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("m", #3/1/1900#): out = out & "DatePart.3_1_1900.m=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("yyyy", 13, #3/1/1900#)): out = out & "DateAdd.3_1_1900.yyyy=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("yyyy", #3/1/1900#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.3_1_1900.yyyy=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("yyyy", #3/1/1900#): out = out & "DatePart.3_1_1900.yyyy=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("h", 13, #3/1/1900#)): out = out & "DateAdd.3_1_1900.h=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("h", #3/1/1900#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.3_1_1900.h=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("h", #3/1/1900#): out = out & "DatePart.3_1_1900.h=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("n", 13, #3/1/1900#)): out = out & "DateAdd.3_1_1900.n=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("n", #3/1/1900#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.3_1_1900.n=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("n", #3/1/1900#): out = out & "DatePart.3_1_1900.n=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("s", 13, #3/1/1900#)): out = out & "DateAdd.3_1_1900.s=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("s", #3/1/1900#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.3_1_1900.s=" & Rz(Err.Number, r) & vbLf
  L015 = out
End Function
Function L016() As String
  Dim r As Variant, out As String, vA0 As Variant, vA1 As Variant, vA2 As Variant, vA3 As Variant, vA4 As Variant, vA5 As Variant, vA6 As Variant, vA7 As Variant, vA8 As Variant, vA9 As Variant, vA10 As Variant, vA11 As Variant, vD0 As Variant, vD1 As Variant, vD2 As Variant, vD3 As Variant, vD4 As Variant, vD5 As Variant, vD6 As Variant, vD7 As Variant, vD8 As Variant, vD9 As Variant, vD10 As Variant, vD11 As Variant, vD12 As Variant, vD13 As Variant, vD14 As Variant, vD15 As Variant, vD16 As Variant, vD17 As Variant
  On Error Resume Next
  vA0 = "": vA1 = "a": vA2 = "abc": vA3 = "Hello World": vA4 = "  pad  ": vA5 = "12": vA6 = "-3.5": vA7 = "1e3": vA8 = "&H1F": vA9 = "abc abc": vA10 = Null: vA11 = Empty
  vD0 = 0: vD1 = 1: vD2 = -1: vD3 = 0.5: vD4 = -0.5: vD5 = 1.5: vD6 = 2.5: vD7 = -1.5: vD8 = -2.5: vD9 = 3.5: vD10 = 100.5: vD11 = 255.5: vD12 = 32767.5: vD13 = 2147483647.5: vD14 = 1E+10: vD15 = 1E+20: vD16 = 123456.789: vD17 = -123456.789
  Err.Clear: r = Empty: r = DatePart("s", #3/1/1900#): out = out & "DatePart.3_1_1900.s=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("w", 13, #3/1/1900#)): out = out & "DateAdd.3_1_1900.w=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("w", #3/1/1900#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.3_1_1900.w=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("w", #3/1/1900#): out = out & "DatePart.3_1_1900.w=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("ww", 13, #3/1/1900#)): out = out & "DateAdd.3_1_1900.ww=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("ww", #3/1/1900#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.3_1_1900.ww=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("ww", #3/1/1900#): out = out & "DatePart.3_1_1900.ww=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = CDbl(DateAdd("q", 13, #3/1/1900#)): out = out & "DateAdd.3_1_1900.q=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DateDiff("q", #3/1/1900#, #6/15/2005 8:30:00 AM#): out = out & "DateDiff.3_1_1900.q=" & Rz(Err.Number, r) & vbLf
  Err.Clear: r = Empty: r = DatePart("q", #3/1/1900#): out = out & "DatePart.3_1_1900.q=" & Rz(Err.Number, r) & vbLf
  L016 = out
End Function
Public Function RunAll() As String
  Dim s As String
  s = s & L000()
  s = s & L001()
  s = s & L002()
  s = s & L003()
  s = s & L004()
  s = s & L005()
  s = s & L006()
  s = s & L007()
  s = s & L008()
  s = s & L009()
  s = s & L010()
  s = s & L011()
  s = s & L012()
  s = s & L013()
  s = s & L014()
  s = s & L015()
  s = s & L016()
  RunAll = s
End Function
