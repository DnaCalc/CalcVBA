Function Rz(e As Long, val As Variant) As String
  If e <> 0 Then
    Rz = "ERR" & e
  Else
    Rz = TypeName(val) & ":" & Trim$(Str$(val))
  End If
End Function
Function P000() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: bA = 0: bB = 0: r = bA + bB: out = out & "bin.Byte0.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: bB = 0: r = bA - bB: out = out & "bin.Byte0.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: bB = 0: r = bA * bB: out = out & "bin.Byte0.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: bB = 0: r = bA / bB: out = out & "bin.Byte0.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: bB = 0: r = bA \ bB: out = out & "bin.Byte0.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: bB = 0: r = bA Mod bB: out = out & "bin.Byte0.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: bB = 0: r = bA ^ bB: out = out & "bin.Byte0.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: bB = 200: r = bA + bB: out = out & "bin.Byte0.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: bB = 200: r = bA - bB: out = out & "bin.Byte0.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: bB = 200: r = bA * bB: out = out & "bin.Byte0.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: bB = 200: r = bA / bB: out = out & "bin.Byte0.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: bB = 200: r = bA \ bB: out = out & "bin.Byte0.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: bB = 200: r = bA Mod bB: out = out & "bin.Byte0.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: bB = 200: r = bA ^ bB: out = out & "bin.Byte0.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: bB = 255: r = bA + bB: out = out & "bin.Byte0.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: bB = 255: r = bA - bB: out = out & "bin.Byte0.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: bB = 255: r = bA * bB: out = out & "bin.Byte0.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: bB = 255: r = bA / bB: out = out & "bin.Byte0.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: bB = 255: r = bA \ bB: out = out & "bin.Byte0.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: bB = 255: r = bA Mod bB: out = out & "bin.Byte0.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: bB = 255: r = bA ^ bB: out = out & "bin.Byte0.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: bB = 0: r = bA + bB: out = out & "bin.Byte1.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: bB = 0: r = bA - bB: out = out & "bin.Byte1.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: bB = 0: r = bA * bB: out = out & "bin.Byte1.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: bB = 0: r = bA / bB: out = out & "bin.Byte1.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: bB = 0: r = bA \ bB: out = out & "bin.Byte1.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: bB = 0: r = bA Mod bB: out = out & "bin.Byte1.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: bB = 0: r = bA ^ bB: out = out & "bin.Byte1.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: bB = 200: r = bA + bB: out = out & "bin.Byte1.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: bB = 200: r = bA - bB: out = out & "bin.Byte1.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: bB = 200: r = bA * bB: out = out & "bin.Byte1.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: bB = 200: r = bA / bB: out = out & "bin.Byte1.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: bB = 200: r = bA \ bB: out = out & "bin.Byte1.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: bB = 200: r = bA Mod bB: out = out & "bin.Byte1.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: bB = 200: r = bA ^ bB: out = out & "bin.Byte1.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: bB = 255: r = bA + bB: out = out & "bin.Byte1.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: bB = 255: r = bA - bB: out = out & "bin.Byte1.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: bB = 255: r = bA * bB: out = out & "bin.Byte1.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: bB = 255: r = bA / bB: out = out & "bin.Byte1.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: bB = 255: r = bA \ bB: out = out & "bin.Byte1.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: bB = 255: r = bA Mod bB: out = out & "bin.Byte1.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: bB = 255: r = bA ^ bB: out = out & "bin.Byte1.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: bB = 0: r = bA + bB: out = out & "bin.Byte2.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: bB = 0: r = bA - bB: out = out & "bin.Byte2.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: bB = 0: r = bA * bB: out = out & "bin.Byte2.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: bB = 0: r = bA / bB: out = out & "bin.Byte2.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: bB = 0: r = bA \ bB: out = out & "bin.Byte2.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: bB = 0: r = bA Mod bB: out = out & "bin.Byte2.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: bB = 0: r = bA ^ bB: out = out & "bin.Byte2.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: bB = 200: r = bA + bB: out = out & "bin.Byte2.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: bB = 200: r = bA - bB: out = out & "bin.Byte2.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: bB = 200: r = bA * bB: out = out & "bin.Byte2.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: bB = 200: r = bA / bB: out = out & "bin.Byte2.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: bB = 200: r = bA \ bB: out = out & "bin.Byte2.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: bB = 200: r = bA Mod bB: out = out & "bin.Byte2.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: bB = 200: r = bA ^ bB: out = out & "bin.Byte2.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: bB = 255: r = bA + bB: out = out & "bin.Byte2.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: bB = 255: r = bA - bB: out = out & "bin.Byte2.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: bB = 255: r = bA * bB: out = out & "bin.Byte2.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: bB = 255: r = bA / bB: out = out & "bin.Byte2.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: bB = 255: r = bA \ bB: out = out & "bin.Byte2.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: bB = 255: r = bA Mod bB: out = out & "bin.Byte2.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: bB = 255: r = bA ^ bB: out = out & "bin.Byte2.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: iB = -32768: r = bA + iB: out = out & "bin.Byte0.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: iB = -32768: r = bA - iB: out = out & "bin.Byte0.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: iB = -32768: r = bA * iB: out = out & "bin.Byte0.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: iB = -32768: r = bA / iB: out = out & "bin.Byte0.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: iB = -32768: r = bA \ iB: out = out & "bin.Byte0.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: iB = -32768: r = bA Mod iB: out = out & "bin.Byte0.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: iB = -32768: r = bA ^ iB: out = out & "bin.Byte0.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: iB = 1: r = bA + iB: out = out & "bin.Byte0.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: iB = 1: r = bA - iB: out = out & "bin.Byte0.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: iB = 1: r = bA * iB: out = out & "bin.Byte0.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: iB = 1: r = bA / iB: out = out & "bin.Byte0.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: iB = 1: r = bA \ iB: out = out & "bin.Byte0.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: iB = 1: r = bA Mod iB: out = out & "bin.Byte0.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: iB = 1: r = bA ^ iB: out = out & "bin.Byte0.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: iB = 32767: r = bA + iB: out = out & "bin.Byte0.add.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: iB = 32767: r = bA - iB: out = out & "bin.Byte0.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: iB = 32767: r = bA * iB: out = out & "bin.Byte0.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: iB = 32767: r = bA / iB: out = out & "bin.Byte0.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: iB = 32767: r = bA \ iB: out = out & "bin.Byte0.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: iB = 32767: r = bA Mod iB: out = out & "bin.Byte0.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: iB = 32767: r = bA ^ iB: out = out & "bin.Byte0.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: iB = -32768: r = bA + iB: out = out & "bin.Byte1.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: iB = -32768: r = bA - iB: out = out & "bin.Byte1.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: iB = -32768: r = bA * iB: out = out & "bin.Byte1.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: iB = -32768: r = bA / iB: out = out & "bin.Byte1.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: iB = -32768: r = bA \ iB: out = out & "bin.Byte1.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: iB = -32768: r = bA Mod iB: out = out & "bin.Byte1.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: iB = -32768: r = bA ^ iB: out = out & "bin.Byte1.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: iB = 1: r = bA + iB: out = out & "bin.Byte1.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: iB = 1: r = bA - iB: out = out & "bin.Byte1.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: iB = 1: r = bA * iB: out = out & "bin.Byte1.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: iB = 1: r = bA / iB: out = out & "bin.Byte1.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: iB = 1: r = bA \ iB: out = out & "bin.Byte1.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: iB = 1: r = bA Mod iB: out = out & "bin.Byte1.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: iB = 1: r = bA ^ iB: out = out & "bin.Byte1.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: iB = 32767: r = bA + iB: out = out & "bin.Byte1.add.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: iB = 32767: r = bA - iB: out = out & "bin.Byte1.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: iB = 32767: r = bA * iB: out = out & "bin.Byte1.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: iB = 32767: r = bA / iB: out = out & "bin.Byte1.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: iB = 32767: r = bA \ iB: out = out & "bin.Byte1.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: iB = 32767: r = bA Mod iB: out = out & "bin.Byte1.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: iB = 32767: r = bA ^ iB: out = out & "bin.Byte1.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: iB = -32768: r = bA + iB: out = out & "bin.Byte2.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: iB = -32768: r = bA - iB: out = out & "bin.Byte2.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: iB = -32768: r = bA * iB: out = out & "bin.Byte2.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: iB = -32768: r = bA / iB: out = out & "bin.Byte2.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: iB = -32768: r = bA \ iB: out = out & "bin.Byte2.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: iB = -32768: r = bA Mod iB: out = out & "bin.Byte2.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: iB = -32768: r = bA ^ iB: out = out & "bin.Byte2.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: iB = 1: r = bA + iB: out = out & "bin.Byte2.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: iB = 1: r = bA - iB: out = out & "bin.Byte2.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: iB = 1: r = bA * iB: out = out & "bin.Byte2.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: iB = 1: r = bA / iB: out = out & "bin.Byte2.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: iB = 1: r = bA \ iB: out = out & "bin.Byte2.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: iB = 1: r = bA Mod iB: out = out & "bin.Byte2.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: iB = 1: r = bA ^ iB: out = out & "bin.Byte2.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: iB = 32767: r = bA + iB: out = out & "bin.Byte2.add.Integer2=" & Rz(Err.Number, r) & vbLf
  P000 = out
End Function
Function P001() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: bA = 255: iB = 32767: r = bA - iB: out = out & "bin.Byte2.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: iB = 32767: r = bA * iB: out = out & "bin.Byte2.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: iB = 32767: r = bA / iB: out = out & "bin.Byte2.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: iB = 32767: r = bA \ iB: out = out & "bin.Byte2.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: iB = 32767: r = bA Mod iB: out = out & "bin.Byte2.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: iB = 32767: r = bA ^ iB: out = out & "bin.Byte2.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: lB = -2147483648: r = bA + lB: out = out & "bin.Byte0.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: lB = -2147483648: r = bA - lB: out = out & "bin.Byte0.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: lB = -2147483648: r = bA * lB: out = out & "bin.Byte0.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: lB = -2147483648: r = bA / lB: out = out & "bin.Byte0.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: lB = -2147483648: r = bA \ lB: out = out & "bin.Byte0.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: lB = -2147483648: r = bA Mod lB: out = out & "bin.Byte0.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: lB = -2147483648: r = bA ^ lB: out = out & "bin.Byte0.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: lB = 65536: r = bA + lB: out = out & "bin.Byte0.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: lB = 65536: r = bA - lB: out = out & "bin.Byte0.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: lB = 65536: r = bA * lB: out = out & "bin.Byte0.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: lB = 65536: r = bA / lB: out = out & "bin.Byte0.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: lB = 65536: r = bA \ lB: out = out & "bin.Byte0.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: lB = 65536: r = bA Mod lB: out = out & "bin.Byte0.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: lB = 65536: r = bA ^ lB: out = out & "bin.Byte0.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: lB = 2147483647: r = bA + lB: out = out & "bin.Byte0.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: lB = 2147483647: r = bA - lB: out = out & "bin.Byte0.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: lB = 2147483647: r = bA * lB: out = out & "bin.Byte0.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: lB = 2147483647: r = bA / lB: out = out & "bin.Byte0.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: lB = 2147483647: r = bA \ lB: out = out & "bin.Byte0.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: lB = 2147483647: r = bA Mod lB: out = out & "bin.Byte0.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: lB = 2147483647: r = bA ^ lB: out = out & "bin.Byte0.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: lB = -2147483648: r = bA + lB: out = out & "bin.Byte1.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: lB = -2147483648: r = bA - lB: out = out & "bin.Byte1.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: lB = -2147483648: r = bA * lB: out = out & "bin.Byte1.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: lB = -2147483648: r = bA / lB: out = out & "bin.Byte1.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: lB = -2147483648: r = bA \ lB: out = out & "bin.Byte1.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: lB = -2147483648: r = bA Mod lB: out = out & "bin.Byte1.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: lB = -2147483648: r = bA ^ lB: out = out & "bin.Byte1.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: lB = 65536: r = bA + lB: out = out & "bin.Byte1.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: lB = 65536: r = bA - lB: out = out & "bin.Byte1.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: lB = 65536: r = bA * lB: out = out & "bin.Byte1.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: lB = 65536: r = bA / lB: out = out & "bin.Byte1.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: lB = 65536: r = bA \ lB: out = out & "bin.Byte1.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: lB = 65536: r = bA Mod lB: out = out & "bin.Byte1.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: lB = 65536: r = bA ^ lB: out = out & "bin.Byte1.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: lB = 2147483647: r = bA + lB: out = out & "bin.Byte1.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: lB = 2147483647: r = bA - lB: out = out & "bin.Byte1.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: lB = 2147483647: r = bA * lB: out = out & "bin.Byte1.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: lB = 2147483647: r = bA / lB: out = out & "bin.Byte1.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: lB = 2147483647: r = bA \ lB: out = out & "bin.Byte1.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: lB = 2147483647: r = bA Mod lB: out = out & "bin.Byte1.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: lB = 2147483647: r = bA ^ lB: out = out & "bin.Byte1.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: lB = -2147483648: r = bA + lB: out = out & "bin.Byte2.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: lB = -2147483648: r = bA - lB: out = out & "bin.Byte2.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: lB = -2147483648: r = bA * lB: out = out & "bin.Byte2.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: lB = -2147483648: r = bA / lB: out = out & "bin.Byte2.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: lB = -2147483648: r = bA \ lB: out = out & "bin.Byte2.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: lB = -2147483648: r = bA Mod lB: out = out & "bin.Byte2.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: lB = -2147483648: r = bA ^ lB: out = out & "bin.Byte2.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: lB = 65536: r = bA + lB: out = out & "bin.Byte2.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: lB = 65536: r = bA - lB: out = out & "bin.Byte2.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: lB = 65536: r = bA * lB: out = out & "bin.Byte2.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: lB = 65536: r = bA / lB: out = out & "bin.Byte2.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: lB = 65536: r = bA \ lB: out = out & "bin.Byte2.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: lB = 65536: r = bA Mod lB: out = out & "bin.Byte2.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: lB = 65536: r = bA ^ lB: out = out & "bin.Byte2.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: lB = 2147483647: r = bA + lB: out = out & "bin.Byte2.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: lB = 2147483647: r = bA - lB: out = out & "bin.Byte2.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: lB = 2147483647: r = bA * lB: out = out & "bin.Byte2.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: lB = 2147483647: r = bA / lB: out = out & "bin.Byte2.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: lB = 2147483647: r = bA \ lB: out = out & "bin.Byte2.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: lB = 2147483647: r = bA Mod lB: out = out & "bin.Byte2.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: lB = 2147483647: r = bA ^ lB: out = out & "bin.Byte2.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: qB = -9223372036854775807^: r = bA + qB: out = out & "bin.Byte0.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: qB = -9223372036854775807^: r = bA - qB: out = out & "bin.Byte0.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: qB = -9223372036854775807^: r = bA * qB: out = out & "bin.Byte0.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: qB = -9223372036854775807^: r = bA / qB: out = out & "bin.Byte0.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: qB = -9223372036854775807^: r = bA \ qB: out = out & "bin.Byte0.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: qB = -9223372036854775807^: r = bA Mod qB: out = out & "bin.Byte0.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: qB = -9223372036854775807^: r = bA ^ qB: out = out & "bin.Byte0.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: qB = 3: r = bA + qB: out = out & "bin.Byte0.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: qB = 3: r = bA - qB: out = out & "bin.Byte0.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: qB = 3: r = bA * qB: out = out & "bin.Byte0.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: qB = 3: r = bA / qB: out = out & "bin.Byte0.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: qB = 3: r = bA \ qB: out = out & "bin.Byte0.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: qB = 3: r = bA Mod qB: out = out & "bin.Byte0.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: qB = 3: r = bA ^ qB: out = out & "bin.Byte0.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: qB = 9223372036854775807^: r = bA + qB: out = out & "bin.Byte0.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: qB = 9223372036854775807^: r = bA - qB: out = out & "bin.Byte0.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: qB = 9223372036854775807^: r = bA * qB: out = out & "bin.Byte0.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: qB = 9223372036854775807^: r = bA / qB: out = out & "bin.Byte0.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: qB = 9223372036854775807^: r = bA \ qB: out = out & "bin.Byte0.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: qB = 9223372036854775807^: r = bA Mod qB: out = out & "bin.Byte0.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: qB = 9223372036854775807^: r = bA ^ qB: out = out & "bin.Byte0.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: qB = -9223372036854775807^: r = bA + qB: out = out & "bin.Byte1.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: qB = -9223372036854775807^: r = bA - qB: out = out & "bin.Byte1.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: qB = -9223372036854775807^: r = bA * qB: out = out & "bin.Byte1.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: qB = -9223372036854775807^: r = bA / qB: out = out & "bin.Byte1.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: qB = -9223372036854775807^: r = bA \ qB: out = out & "bin.Byte1.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: qB = -9223372036854775807^: r = bA Mod qB: out = out & "bin.Byte1.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: qB = -9223372036854775807^: r = bA ^ qB: out = out & "bin.Byte1.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: qB = 3: r = bA + qB: out = out & "bin.Byte1.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: qB = 3: r = bA - qB: out = out & "bin.Byte1.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: qB = 3: r = bA * qB: out = out & "bin.Byte1.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: qB = 3: r = bA / qB: out = out & "bin.Byte1.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: qB = 3: r = bA \ qB: out = out & "bin.Byte1.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: qB = 3: r = bA Mod qB: out = out & "bin.Byte1.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: qB = 3: r = bA ^ qB: out = out & "bin.Byte1.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: qB = 9223372036854775807^: r = bA + qB: out = out & "bin.Byte1.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: qB = 9223372036854775807^: r = bA - qB: out = out & "bin.Byte1.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: qB = 9223372036854775807^: r = bA * qB: out = out & "bin.Byte1.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: qB = 9223372036854775807^: r = bA / qB: out = out & "bin.Byte1.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: qB = 9223372036854775807^: r = bA \ qB: out = out & "bin.Byte1.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: qB = 9223372036854775807^: r = bA Mod qB: out = out & "bin.Byte1.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: qB = 9223372036854775807^: r = bA ^ qB: out = out & "bin.Byte1.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: qB = -9223372036854775807^: r = bA + qB: out = out & "bin.Byte2.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: qB = -9223372036854775807^: r = bA - qB: out = out & "bin.Byte2.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: qB = -9223372036854775807^: r = bA * qB: out = out & "bin.Byte2.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: qB = -9223372036854775807^: r = bA / qB: out = out & "bin.Byte2.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: qB = -9223372036854775807^: r = bA \ qB: out = out & "bin.Byte2.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: qB = -9223372036854775807^: r = bA Mod qB: out = out & "bin.Byte2.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: qB = -9223372036854775807^: r = bA ^ qB: out = out & "bin.Byte2.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: qB = 3: r = bA + qB: out = out & "bin.Byte2.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: qB = 3: r = bA - qB: out = out & "bin.Byte2.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  P001 = out
End Function
Function P002() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: bA = 255: qB = 3: r = bA * qB: out = out & "bin.Byte2.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: qB = 3: r = bA / qB: out = out & "bin.Byte2.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: qB = 3: r = bA \ qB: out = out & "bin.Byte2.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: qB = 3: r = bA Mod qB: out = out & "bin.Byte2.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: qB = 3: r = bA ^ qB: out = out & "bin.Byte2.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: qB = 9223372036854775807^: r = bA + qB: out = out & "bin.Byte2.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: qB = 9223372036854775807^: r = bA - qB: out = out & "bin.Byte2.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: qB = 9223372036854775807^: r = bA * qB: out = out & "bin.Byte2.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: qB = 9223372036854775807^: r = bA / qB: out = out & "bin.Byte2.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: qB = 9223372036854775807^: r = bA \ qB: out = out & "bin.Byte2.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: qB = 9223372036854775807^: r = bA Mod qB: out = out & "bin.Byte2.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: qB = 9223372036854775807^: r = bA ^ qB: out = out & "bin.Byte2.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: sB = 0.5: r = bA + sB: out = out & "bin.Byte0.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: sB = 0.5: r = bA - sB: out = out & "bin.Byte0.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: sB = 0.5: r = bA * sB: out = out & "bin.Byte0.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: sB = 0.5: r = bA / sB: out = out & "bin.Byte0.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: sB = 0.5: r = bA \ sB: out = out & "bin.Byte0.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: sB = 0.5: r = bA Mod sB: out = out & "bin.Byte0.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: sB = 0.5: r = bA ^ sB: out = out & "bin.Byte0.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: sB = 16777216: r = bA + sB: out = out & "bin.Byte0.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: sB = 16777216: r = bA - sB: out = out & "bin.Byte0.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: sB = 16777216: r = bA * sB: out = out & "bin.Byte0.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: sB = 16777216: r = bA / sB: out = out & "bin.Byte0.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: sB = 16777216: r = bA \ sB: out = out & "bin.Byte0.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: sB = 16777216: r = bA Mod sB: out = out & "bin.Byte0.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: sB = 16777216: r = bA ^ sB: out = out & "bin.Byte0.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: sB = 3.0E+38: r = bA + sB: out = out & "bin.Byte0.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: sB = 3.0E+38: r = bA - sB: out = out & "bin.Byte0.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: sB = 3.0E+38: r = bA * sB: out = out & "bin.Byte0.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: sB = 3.0E+38: r = bA / sB: out = out & "bin.Byte0.div.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: sB = 3.0E+38: r = bA \ sB: out = out & "bin.Byte0.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: sB = 3.0E+38: r = bA Mod sB: out = out & "bin.Byte0.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: sB = 3.0E+38: r = bA ^ sB: out = out & "bin.Byte0.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: sB = 0.5: r = bA + sB: out = out & "bin.Byte1.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: sB = 0.5: r = bA - sB: out = out & "bin.Byte1.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: sB = 0.5: r = bA * sB: out = out & "bin.Byte1.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: sB = 0.5: r = bA / sB: out = out & "bin.Byte1.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: sB = 0.5: r = bA \ sB: out = out & "bin.Byte1.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: sB = 0.5: r = bA Mod sB: out = out & "bin.Byte1.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: sB = 0.5: r = bA ^ sB: out = out & "bin.Byte1.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: sB = 16777216: r = bA + sB: out = out & "bin.Byte1.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: sB = 16777216: r = bA - sB: out = out & "bin.Byte1.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: sB = 16777216: r = bA * sB: out = out & "bin.Byte1.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: sB = 16777216: r = bA / sB: out = out & "bin.Byte1.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: sB = 16777216: r = bA \ sB: out = out & "bin.Byte1.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: sB = 16777216: r = bA Mod sB: out = out & "bin.Byte1.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: sB = 16777216: r = bA ^ sB: out = out & "bin.Byte1.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: sB = 3.0E+38: r = bA + sB: out = out & "bin.Byte1.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: sB = 3.0E+38: r = bA - sB: out = out & "bin.Byte1.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: sB = 3.0E+38: r = bA * sB: out = out & "bin.Byte1.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: sB = 3.0E+38: r = bA / sB: out = out & "bin.Byte1.div.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: sB = 3.0E+38: r = bA \ sB: out = out & "bin.Byte1.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: sB = 3.0E+38: r = bA Mod sB: out = out & "bin.Byte1.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: sB = 3.0E+38: r = bA ^ sB: out = out & "bin.Byte1.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: sB = 0.5: r = bA + sB: out = out & "bin.Byte2.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: sB = 0.5: r = bA - sB: out = out & "bin.Byte2.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: sB = 0.5: r = bA * sB: out = out & "bin.Byte2.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: sB = 0.5: r = bA / sB: out = out & "bin.Byte2.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: sB = 0.5: r = bA \ sB: out = out & "bin.Byte2.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: sB = 0.5: r = bA Mod sB: out = out & "bin.Byte2.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: sB = 0.5: r = bA ^ sB: out = out & "bin.Byte2.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: sB = 16777216: r = bA + sB: out = out & "bin.Byte2.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: sB = 16777216: r = bA - sB: out = out & "bin.Byte2.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: sB = 16777216: r = bA * sB: out = out & "bin.Byte2.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: sB = 16777216: r = bA / sB: out = out & "bin.Byte2.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: sB = 16777216: r = bA \ sB: out = out & "bin.Byte2.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: sB = 16777216: r = bA Mod sB: out = out & "bin.Byte2.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: sB = 16777216: r = bA ^ sB: out = out & "bin.Byte2.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: sB = 3.0E+38: r = bA + sB: out = out & "bin.Byte2.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: sB = 3.0E+38: r = bA - sB: out = out & "bin.Byte2.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: sB = 3.0E+38: r = bA * sB: out = out & "bin.Byte2.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: sB = 3.0E+38: r = bA / sB: out = out & "bin.Byte2.div.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: sB = 3.0E+38: r = bA \ sB: out = out & "bin.Byte2.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: sB = 3.0E+38: r = bA Mod sB: out = out & "bin.Byte2.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: sB = 3.0E+38: r = bA ^ sB: out = out & "bin.Byte2.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: dB = 0.5: r = bA + dB: out = out & "bin.Byte0.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: dB = 0.5: r = bA - dB: out = out & "bin.Byte0.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: dB = 0.5: r = bA * dB: out = out & "bin.Byte0.mul.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: dB = 0.5: r = bA / dB: out = out & "bin.Byte0.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: dB = 0.5: r = bA \ dB: out = out & "bin.Byte0.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: dB = 0.5: r = bA Mod dB: out = out & "bin.Byte0.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: dB = 0.5: r = bA ^ dB: out = out & "bin.Byte0.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: dB = 2.5: r = bA + dB: out = out & "bin.Byte0.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: dB = 2.5: r = bA - dB: out = out & "bin.Byte0.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: dB = 2.5: r = bA * dB: out = out & "bin.Byte0.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: dB = 2.5: r = bA / dB: out = out & "bin.Byte0.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: dB = 2.5: r = bA \ dB: out = out & "bin.Byte0.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: dB = 2.5: r = bA Mod dB: out = out & "bin.Byte0.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: dB = 2.5: r = bA ^ dB: out = out & "bin.Byte0.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: dB = 1.0E+308: r = bA + dB: out = out & "bin.Byte0.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: dB = 1.0E+308: r = bA - dB: out = out & "bin.Byte0.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: dB = 1.0E+308: r = bA * dB: out = out & "bin.Byte0.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: dB = 1.0E+308: r = bA / dB: out = out & "bin.Byte0.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: dB = 1.0E+308: r = bA \ dB: out = out & "bin.Byte0.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: dB = 1.0E+308: r = bA Mod dB: out = out & "bin.Byte0.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: dB = 1.0E+308: r = bA ^ dB: out = out & "bin.Byte0.pow.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: dB = 0.5: r = bA + dB: out = out & "bin.Byte1.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: dB = 0.5: r = bA - dB: out = out & "bin.Byte1.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: dB = 0.5: r = bA * dB: out = out & "bin.Byte1.mul.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: dB = 0.5: r = bA / dB: out = out & "bin.Byte1.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: dB = 0.5: r = bA \ dB: out = out & "bin.Byte1.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: dB = 0.5: r = bA Mod dB: out = out & "bin.Byte1.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: dB = 0.5: r = bA ^ dB: out = out & "bin.Byte1.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: dB = 2.5: r = bA + dB: out = out & "bin.Byte1.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: dB = 2.5: r = bA - dB: out = out & "bin.Byte1.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: dB = 2.5: r = bA * dB: out = out & "bin.Byte1.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: dB = 2.5: r = bA / dB: out = out & "bin.Byte1.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: dB = 2.5: r = bA \ dB: out = out & "bin.Byte1.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: dB = 2.5: r = bA Mod dB: out = out & "bin.Byte1.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: dB = 2.5: r = bA ^ dB: out = out & "bin.Byte1.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: dB = 1.0E+308: r = bA + dB: out = out & "bin.Byte1.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: dB = 1.0E+308: r = bA - dB: out = out & "bin.Byte1.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: dB = 1.0E+308: r = bA * dB: out = out & "bin.Byte1.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: dB = 1.0E+308: r = bA / dB: out = out & "bin.Byte1.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: dB = 1.0E+308: r = bA \ dB: out = out & "bin.Byte1.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: dB = 1.0E+308: r = bA Mod dB: out = out & "bin.Byte1.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: dB = 1.0E+308: r = bA ^ dB: out = out & "bin.Byte1.pow.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: dB = 0.5: r = bA + dB: out = out & "bin.Byte2.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: dB = 0.5: r = bA - dB: out = out & "bin.Byte2.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: dB = 0.5: r = bA * dB: out = out & "bin.Byte2.mul.Double0=" & Rz(Err.Number, r) & vbLf
  P002 = out
End Function
Function P003() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: bA = 255: dB = 0.5: r = bA / dB: out = out & "bin.Byte2.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: dB = 0.5: r = bA \ dB: out = out & "bin.Byte2.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: dB = 0.5: r = bA Mod dB: out = out & "bin.Byte2.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: dB = 0.5: r = bA ^ dB: out = out & "bin.Byte2.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: dB = 2.5: r = bA + dB: out = out & "bin.Byte2.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: dB = 2.5: r = bA - dB: out = out & "bin.Byte2.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: dB = 2.5: r = bA * dB: out = out & "bin.Byte2.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: dB = 2.5: r = bA / dB: out = out & "bin.Byte2.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: dB = 2.5: r = bA \ dB: out = out & "bin.Byte2.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: dB = 2.5: r = bA Mod dB: out = out & "bin.Byte2.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: dB = 2.5: r = bA ^ dB: out = out & "bin.Byte2.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: dB = 1.0E+308: r = bA + dB: out = out & "bin.Byte2.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: dB = 1.0E+308: r = bA - dB: out = out & "bin.Byte2.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: dB = 1.0E+308: r = bA * dB: out = out & "bin.Byte2.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: dB = 1.0E+308: r = bA / dB: out = out & "bin.Byte2.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: dB = 1.0E+308: r = bA \ dB: out = out & "bin.Byte2.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: dB = 1.0E+308: r = bA Mod dB: out = out & "bin.Byte2.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: dB = 1.0E+308: r = bA ^ dB: out = out & "bin.Byte2.pow.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: cB = -0.5@: r = bA + cB: out = out & "bin.Byte0.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: cB = -0.5@: r = bA - cB: out = out & "bin.Byte0.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: cB = -0.5@: r = bA * cB: out = out & "bin.Byte0.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: cB = -0.5@: r = bA / cB: out = out & "bin.Byte0.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: cB = -0.5@: r = bA \ cB: out = out & "bin.Byte0.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: cB = -0.5@: r = bA Mod cB: out = out & "bin.Byte0.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: cB = -0.5@: r = bA ^ cB: out = out & "bin.Byte0.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: cB = 1.5@: r = bA + cB: out = out & "bin.Byte0.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: cB = 1.5@: r = bA - cB: out = out & "bin.Byte0.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: cB = 1.5@: r = bA * cB: out = out & "bin.Byte0.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: cB = 1.5@: r = bA / cB: out = out & "bin.Byte0.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: cB = 1.5@: r = bA \ cB: out = out & "bin.Byte0.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: cB = 1.5@: r = bA Mod cB: out = out & "bin.Byte0.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: cB = 1.5@: r = bA ^ cB: out = out & "bin.Byte0.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: cB = 922337203685477.5807@: r = bA + cB: out = out & "bin.Byte0.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: cB = 922337203685477.5807@: r = bA - cB: out = out & "bin.Byte0.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: cB = 922337203685477.5807@: r = bA * cB: out = out & "bin.Byte0.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: cB = 922337203685477.5807@: r = bA / cB: out = out & "bin.Byte0.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: cB = 922337203685477.5807@: r = bA \ cB: out = out & "bin.Byte0.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: cB = 922337203685477.5807@: r = bA Mod cB: out = out & "bin.Byte0.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: cB = 922337203685477.5807@: r = bA ^ cB: out = out & "bin.Byte0.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: cB = -0.5@: r = bA + cB: out = out & "bin.Byte1.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: cB = -0.5@: r = bA - cB: out = out & "bin.Byte1.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: cB = -0.5@: r = bA * cB: out = out & "bin.Byte1.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: cB = -0.5@: r = bA / cB: out = out & "bin.Byte1.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: cB = -0.5@: r = bA \ cB: out = out & "bin.Byte1.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: cB = -0.5@: r = bA Mod cB: out = out & "bin.Byte1.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: cB = -0.5@: r = bA ^ cB: out = out & "bin.Byte1.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: cB = 1.5@: r = bA + cB: out = out & "bin.Byte1.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: cB = 1.5@: r = bA - cB: out = out & "bin.Byte1.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: cB = 1.5@: r = bA * cB: out = out & "bin.Byte1.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: cB = 1.5@: r = bA / cB: out = out & "bin.Byte1.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: cB = 1.5@: r = bA \ cB: out = out & "bin.Byte1.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: cB = 1.5@: r = bA Mod cB: out = out & "bin.Byte1.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: cB = 1.5@: r = bA ^ cB: out = out & "bin.Byte1.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: cB = 922337203685477.5807@: r = bA + cB: out = out & "bin.Byte1.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: cB = 922337203685477.5807@: r = bA - cB: out = out & "bin.Byte1.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: cB = 922337203685477.5807@: r = bA * cB: out = out & "bin.Byte1.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: cB = 922337203685477.5807@: r = bA / cB: out = out & "bin.Byte1.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: cB = 922337203685477.5807@: r = bA \ cB: out = out & "bin.Byte1.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: cB = 922337203685477.5807@: r = bA Mod cB: out = out & "bin.Byte1.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: cB = 922337203685477.5807@: r = bA ^ cB: out = out & "bin.Byte1.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: cB = -0.5@: r = bA + cB: out = out & "bin.Byte2.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: cB = -0.5@: r = bA - cB: out = out & "bin.Byte2.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: cB = -0.5@: r = bA * cB: out = out & "bin.Byte2.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: cB = -0.5@: r = bA / cB: out = out & "bin.Byte2.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: cB = -0.5@: r = bA \ cB: out = out & "bin.Byte2.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: cB = -0.5@: r = bA Mod cB: out = out & "bin.Byte2.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: cB = -0.5@: r = bA ^ cB: out = out & "bin.Byte2.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: cB = 1.5@: r = bA + cB: out = out & "bin.Byte2.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: cB = 1.5@: r = bA - cB: out = out & "bin.Byte2.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: cB = 1.5@: r = bA * cB: out = out & "bin.Byte2.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: cB = 1.5@: r = bA / cB: out = out & "bin.Byte2.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: cB = 1.5@: r = bA \ cB: out = out & "bin.Byte2.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: cB = 1.5@: r = bA Mod cB: out = out & "bin.Byte2.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: cB = 1.5@: r = bA ^ cB: out = out & "bin.Byte2.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: cB = 922337203685477.5807@: r = bA + cB: out = out & "bin.Byte2.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: cB = 922337203685477.5807@: r = bA - cB: out = out & "bin.Byte2.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: cB = 922337203685477.5807@: r = bA * cB: out = out & "bin.Byte2.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: cB = 922337203685477.5807@: r = bA / cB: out = out & "bin.Byte2.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: cB = 922337203685477.5807@: r = bA \ cB: out = out & "bin.Byte2.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: cB = 922337203685477.5807@: r = bA Mod cB: out = out & "bin.Byte2.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: cB = 922337203685477.5807@: r = bA ^ cB: out = out & "bin.Byte2.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: vB = Empty: r = bA + vB: out = out & "bin.Byte0.add.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: vB = Empty: r = bA - vB: out = out & "bin.Byte0.sub.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: vB = Empty: r = bA * vB: out = out & "bin.Byte0.mul.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: vB = Empty: r = bA / vB: out = out & "bin.Byte0.div.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: vB = Empty: r = bA \ vB: out = out & "bin.Byte0.idiv.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: vB = Empty: r = bA Mod vB: out = out & "bin.Byte0.mod.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: vB = Empty: r = bA ^ vB: out = out & "bin.Byte0.pow.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: vB = 7: r = bA + vB: out = out & "bin.Byte0.add.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: vB = 7: r = bA - vB: out = out & "bin.Byte0.sub.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: vB = 7: r = bA * vB: out = out & "bin.Byte0.mul.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: vB = 7: r = bA / vB: out = out & "bin.Byte0.div.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: vB = 7: r = bA \ vB: out = out & "bin.Byte0.idiv.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: vB = 7: r = bA Mod vB: out = out & "bin.Byte0.mod.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: vB = 7: r = bA ^ vB: out = out & "bin.Byte0.pow.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: vB = 2147483647: r = bA + vB: out = out & "bin.Byte0.add.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: vB = 2147483647: r = bA - vB: out = out & "bin.Byte0.sub.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: vB = 2147483647: r = bA * vB: out = out & "bin.Byte0.mul.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: vB = 2147483647: r = bA / vB: out = out & "bin.Byte0.div.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: vB = 2147483647: r = bA \ vB: out = out & "bin.Byte0.idiv.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: vB = 2147483647: r = bA Mod vB: out = out & "bin.Byte0.mod.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: vB = 2147483647: r = bA ^ vB: out = out & "bin.Byte0.pow.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: vB = Empty: r = bA + vB: out = out & "bin.Byte1.add.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: vB = Empty: r = bA - vB: out = out & "bin.Byte1.sub.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: vB = Empty: r = bA * vB: out = out & "bin.Byte1.mul.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: vB = Empty: r = bA / vB: out = out & "bin.Byte1.div.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: vB = Empty: r = bA \ vB: out = out & "bin.Byte1.idiv.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: vB = Empty: r = bA Mod vB: out = out & "bin.Byte1.mod.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: vB = Empty: r = bA ^ vB: out = out & "bin.Byte1.pow.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: vB = 7: r = bA + vB: out = out & "bin.Byte1.add.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: vB = 7: r = bA - vB: out = out & "bin.Byte1.sub.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: vB = 7: r = bA * vB: out = out & "bin.Byte1.mul.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: vB = 7: r = bA / vB: out = out & "bin.Byte1.div.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: vB = 7: r = bA \ vB: out = out & "bin.Byte1.idiv.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: vB = 7: r = bA Mod vB: out = out & "bin.Byte1.mod.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: vB = 7: r = bA ^ vB: out = out & "bin.Byte1.pow.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: vB = 2147483647: r = bA + vB: out = out & "bin.Byte1.add.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: vB = 2147483647: r = bA - vB: out = out & "bin.Byte1.sub.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: vB = 2147483647: r = bA * vB: out = out & "bin.Byte1.mul.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: vB = 2147483647: r = bA / vB: out = out & "bin.Byte1.div.Variant2=" & Rz(Err.Number, r) & vbLf
  P003 = out
End Function
Function P004() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: bA = 200: vB = 2147483647: r = bA \ vB: out = out & "bin.Byte1.idiv.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: vB = 2147483647: r = bA Mod vB: out = out & "bin.Byte1.mod.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: vB = 2147483647: r = bA ^ vB: out = out & "bin.Byte1.pow.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: vB = Empty: r = bA + vB: out = out & "bin.Byte2.add.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: vB = Empty: r = bA - vB: out = out & "bin.Byte2.sub.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: vB = Empty: r = bA * vB: out = out & "bin.Byte2.mul.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: vB = Empty: r = bA / vB: out = out & "bin.Byte2.div.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: vB = Empty: r = bA \ vB: out = out & "bin.Byte2.idiv.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: vB = Empty: r = bA Mod vB: out = out & "bin.Byte2.mod.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: vB = Empty: r = bA ^ vB: out = out & "bin.Byte2.pow.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: vB = 7: r = bA + vB: out = out & "bin.Byte2.add.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: vB = 7: r = bA - vB: out = out & "bin.Byte2.sub.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: vB = 7: r = bA * vB: out = out & "bin.Byte2.mul.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: vB = 7: r = bA / vB: out = out & "bin.Byte2.div.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: vB = 7: r = bA \ vB: out = out & "bin.Byte2.idiv.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: vB = 7: r = bA Mod vB: out = out & "bin.Byte2.mod.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: vB = 7: r = bA ^ vB: out = out & "bin.Byte2.pow.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: vB = 2147483647: r = bA + vB: out = out & "bin.Byte2.add.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: vB = 2147483647: r = bA - vB: out = out & "bin.Byte2.sub.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: vB = 2147483647: r = bA * vB: out = out & "bin.Byte2.mul.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: vB = 2147483647: r = bA / vB: out = out & "bin.Byte2.div.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: vB = 2147483647: r = bA \ vB: out = out & "bin.Byte2.idiv.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: vB = 2147483647: r = bA Mod vB: out = out & "bin.Byte2.mod.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: vB = 2147483647: r = bA ^ vB: out = out & "bin.Byte2.pow.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: bB = 0: r = iA + bB: out = out & "bin.Integer0.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: bB = 0: r = iA - bB: out = out & "bin.Integer0.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: bB = 0: r = iA * bB: out = out & "bin.Integer0.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: bB = 0: r = iA / bB: out = out & "bin.Integer0.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: bB = 0: r = iA \ bB: out = out & "bin.Integer0.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: bB = 0: r = iA Mod bB: out = out & "bin.Integer0.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: bB = 0: r = iA ^ bB: out = out & "bin.Integer0.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: bB = 200: r = iA + bB: out = out & "bin.Integer0.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: bB = 200: r = iA - bB: out = out & "bin.Integer0.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: bB = 200: r = iA * bB: out = out & "bin.Integer0.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: bB = 200: r = iA / bB: out = out & "bin.Integer0.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: bB = 200: r = iA \ bB: out = out & "bin.Integer0.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: bB = 200: r = iA Mod bB: out = out & "bin.Integer0.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: bB = 200: r = iA ^ bB: out = out & "bin.Integer0.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: bB = 255: r = iA + bB: out = out & "bin.Integer0.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: bB = 255: r = iA - bB: out = out & "bin.Integer0.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: bB = 255: r = iA * bB: out = out & "bin.Integer0.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: bB = 255: r = iA / bB: out = out & "bin.Integer0.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: bB = 255: r = iA \ bB: out = out & "bin.Integer0.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: bB = 255: r = iA Mod bB: out = out & "bin.Integer0.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: bB = 255: r = iA ^ bB: out = out & "bin.Integer0.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: bB = 0: r = iA + bB: out = out & "bin.Integer1.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: bB = 0: r = iA - bB: out = out & "bin.Integer1.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: bB = 0: r = iA * bB: out = out & "bin.Integer1.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: bB = 0: r = iA / bB: out = out & "bin.Integer1.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: bB = 0: r = iA \ bB: out = out & "bin.Integer1.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: bB = 0: r = iA Mod bB: out = out & "bin.Integer1.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: bB = 0: r = iA ^ bB: out = out & "bin.Integer1.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: bB = 200: r = iA + bB: out = out & "bin.Integer1.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: bB = 200: r = iA - bB: out = out & "bin.Integer1.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: bB = 200: r = iA * bB: out = out & "bin.Integer1.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: bB = 200: r = iA / bB: out = out & "bin.Integer1.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: bB = 200: r = iA \ bB: out = out & "bin.Integer1.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: bB = 200: r = iA Mod bB: out = out & "bin.Integer1.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: bB = 200: r = iA ^ bB: out = out & "bin.Integer1.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: bB = 255: r = iA + bB: out = out & "bin.Integer1.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: bB = 255: r = iA - bB: out = out & "bin.Integer1.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: bB = 255: r = iA * bB: out = out & "bin.Integer1.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: bB = 255: r = iA / bB: out = out & "bin.Integer1.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: bB = 255: r = iA \ bB: out = out & "bin.Integer1.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: bB = 255: r = iA Mod bB: out = out & "bin.Integer1.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: bB = 255: r = iA ^ bB: out = out & "bin.Integer1.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: bB = 0: r = iA + bB: out = out & "bin.Integer2.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: bB = 0: r = iA - bB: out = out & "bin.Integer2.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: bB = 0: r = iA * bB: out = out & "bin.Integer2.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: bB = 0: r = iA / bB: out = out & "bin.Integer2.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: bB = 0: r = iA \ bB: out = out & "bin.Integer2.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: bB = 0: r = iA Mod bB: out = out & "bin.Integer2.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: bB = 0: r = iA ^ bB: out = out & "bin.Integer2.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: bB = 200: r = iA + bB: out = out & "bin.Integer2.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: bB = 200: r = iA - bB: out = out & "bin.Integer2.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: bB = 200: r = iA * bB: out = out & "bin.Integer2.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: bB = 200: r = iA / bB: out = out & "bin.Integer2.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: bB = 200: r = iA \ bB: out = out & "bin.Integer2.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: bB = 200: r = iA Mod bB: out = out & "bin.Integer2.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: bB = 200: r = iA ^ bB: out = out & "bin.Integer2.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: bB = 255: r = iA + bB: out = out & "bin.Integer2.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: bB = 255: r = iA - bB: out = out & "bin.Integer2.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: bB = 255: r = iA * bB: out = out & "bin.Integer2.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: bB = 255: r = iA / bB: out = out & "bin.Integer2.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: bB = 255: r = iA \ bB: out = out & "bin.Integer2.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: bB = 255: r = iA Mod bB: out = out & "bin.Integer2.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: bB = 255: r = iA ^ bB: out = out & "bin.Integer2.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: iB = -32768: r = iA + iB: out = out & "bin.Integer0.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: iB = -32768: r = iA - iB: out = out & "bin.Integer0.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: iB = -32768: r = iA * iB: out = out & "bin.Integer0.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: iB = -32768: r = iA / iB: out = out & "bin.Integer0.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: iB = -32768: r = iA \ iB: out = out & "bin.Integer0.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: iB = -32768: r = iA Mod iB: out = out & "bin.Integer0.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: iB = -32768: r = iA ^ iB: out = out & "bin.Integer0.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: iB = 1: r = iA + iB: out = out & "bin.Integer0.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: iB = 1: r = iA - iB: out = out & "bin.Integer0.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: iB = 1: r = iA * iB: out = out & "bin.Integer0.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: iB = 1: r = iA / iB: out = out & "bin.Integer0.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: iB = 1: r = iA \ iB: out = out & "bin.Integer0.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: iB = 1: r = iA Mod iB: out = out & "bin.Integer0.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: iB = 1: r = iA ^ iB: out = out & "bin.Integer0.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: iB = 32767: r = iA + iB: out = out & "bin.Integer0.add.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: iB = 32767: r = iA - iB: out = out & "bin.Integer0.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: iB = 32767: r = iA * iB: out = out & "bin.Integer0.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: iB = 32767: r = iA / iB: out = out & "bin.Integer0.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: iB = 32767: r = iA \ iB: out = out & "bin.Integer0.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: iB = 32767: r = iA Mod iB: out = out & "bin.Integer0.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: iB = 32767: r = iA ^ iB: out = out & "bin.Integer0.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: iB = -32768: r = iA + iB: out = out & "bin.Integer1.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: iB = -32768: r = iA - iB: out = out & "bin.Integer1.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: iB = -32768: r = iA * iB: out = out & "bin.Integer1.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: iB = -32768: r = iA / iB: out = out & "bin.Integer1.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: iB = -32768: r = iA \ iB: out = out & "bin.Integer1.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: iB = -32768: r = iA Mod iB: out = out & "bin.Integer1.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: iB = -32768: r = iA ^ iB: out = out & "bin.Integer1.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: iB = 1: r = iA + iB: out = out & "bin.Integer1.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: iB = 1: r = iA - iB: out = out & "bin.Integer1.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: iB = 1: r = iA * iB: out = out & "bin.Integer1.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: iB = 1: r = iA / iB: out = out & "bin.Integer1.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: iB = 1: r = iA \ iB: out = out & "bin.Integer1.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  P004 = out
End Function
Function P005() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: iA = 1: iB = 1: r = iA Mod iB: out = out & "bin.Integer1.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: iB = 1: r = iA ^ iB: out = out & "bin.Integer1.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: iB = 32767: r = iA + iB: out = out & "bin.Integer1.add.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: iB = 32767: r = iA - iB: out = out & "bin.Integer1.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: iB = 32767: r = iA * iB: out = out & "bin.Integer1.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: iB = 32767: r = iA / iB: out = out & "bin.Integer1.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: iB = 32767: r = iA \ iB: out = out & "bin.Integer1.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: iB = 32767: r = iA Mod iB: out = out & "bin.Integer1.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: iB = 32767: r = iA ^ iB: out = out & "bin.Integer1.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: iB = -32768: r = iA + iB: out = out & "bin.Integer2.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: iB = -32768: r = iA - iB: out = out & "bin.Integer2.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: iB = -32768: r = iA * iB: out = out & "bin.Integer2.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: iB = -32768: r = iA / iB: out = out & "bin.Integer2.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: iB = -32768: r = iA \ iB: out = out & "bin.Integer2.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: iB = -32768: r = iA Mod iB: out = out & "bin.Integer2.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: iB = -32768: r = iA ^ iB: out = out & "bin.Integer2.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: iB = 1: r = iA + iB: out = out & "bin.Integer2.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: iB = 1: r = iA - iB: out = out & "bin.Integer2.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: iB = 1: r = iA * iB: out = out & "bin.Integer2.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: iB = 1: r = iA / iB: out = out & "bin.Integer2.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: iB = 1: r = iA \ iB: out = out & "bin.Integer2.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: iB = 1: r = iA Mod iB: out = out & "bin.Integer2.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: iB = 1: r = iA ^ iB: out = out & "bin.Integer2.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: iB = 32767: r = iA + iB: out = out & "bin.Integer2.add.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: iB = 32767: r = iA - iB: out = out & "bin.Integer2.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: iB = 32767: r = iA * iB: out = out & "bin.Integer2.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: iB = 32767: r = iA / iB: out = out & "bin.Integer2.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: iB = 32767: r = iA \ iB: out = out & "bin.Integer2.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: iB = 32767: r = iA Mod iB: out = out & "bin.Integer2.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: iB = 32767: r = iA ^ iB: out = out & "bin.Integer2.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: lB = -2147483648: r = iA + lB: out = out & "bin.Integer0.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: lB = -2147483648: r = iA - lB: out = out & "bin.Integer0.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: lB = -2147483648: r = iA * lB: out = out & "bin.Integer0.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: lB = -2147483648: r = iA / lB: out = out & "bin.Integer0.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: lB = -2147483648: r = iA \ lB: out = out & "bin.Integer0.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: lB = -2147483648: r = iA Mod lB: out = out & "bin.Integer0.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: lB = -2147483648: r = iA ^ lB: out = out & "bin.Integer0.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: lB = 65536: r = iA + lB: out = out & "bin.Integer0.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: lB = 65536: r = iA - lB: out = out & "bin.Integer0.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: lB = 65536: r = iA * lB: out = out & "bin.Integer0.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: lB = 65536: r = iA / lB: out = out & "bin.Integer0.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: lB = 65536: r = iA \ lB: out = out & "bin.Integer0.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: lB = 65536: r = iA Mod lB: out = out & "bin.Integer0.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: lB = 65536: r = iA ^ lB: out = out & "bin.Integer0.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: lB = 2147483647: r = iA + lB: out = out & "bin.Integer0.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: lB = 2147483647: r = iA - lB: out = out & "bin.Integer0.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: lB = 2147483647: r = iA * lB: out = out & "bin.Integer0.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: lB = 2147483647: r = iA / lB: out = out & "bin.Integer0.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: lB = 2147483647: r = iA \ lB: out = out & "bin.Integer0.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: lB = 2147483647: r = iA Mod lB: out = out & "bin.Integer0.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: lB = 2147483647: r = iA ^ lB: out = out & "bin.Integer0.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: lB = -2147483648: r = iA + lB: out = out & "bin.Integer1.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: lB = -2147483648: r = iA - lB: out = out & "bin.Integer1.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: lB = -2147483648: r = iA * lB: out = out & "bin.Integer1.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: lB = -2147483648: r = iA / lB: out = out & "bin.Integer1.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: lB = -2147483648: r = iA \ lB: out = out & "bin.Integer1.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: lB = -2147483648: r = iA Mod lB: out = out & "bin.Integer1.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: lB = -2147483648: r = iA ^ lB: out = out & "bin.Integer1.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: lB = 65536: r = iA + lB: out = out & "bin.Integer1.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: lB = 65536: r = iA - lB: out = out & "bin.Integer1.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: lB = 65536: r = iA * lB: out = out & "bin.Integer1.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: lB = 65536: r = iA / lB: out = out & "bin.Integer1.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: lB = 65536: r = iA \ lB: out = out & "bin.Integer1.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: lB = 65536: r = iA Mod lB: out = out & "bin.Integer1.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: lB = 65536: r = iA ^ lB: out = out & "bin.Integer1.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: lB = 2147483647: r = iA + lB: out = out & "bin.Integer1.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: lB = 2147483647: r = iA - lB: out = out & "bin.Integer1.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: lB = 2147483647: r = iA * lB: out = out & "bin.Integer1.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: lB = 2147483647: r = iA / lB: out = out & "bin.Integer1.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: lB = 2147483647: r = iA \ lB: out = out & "bin.Integer1.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: lB = 2147483647: r = iA Mod lB: out = out & "bin.Integer1.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: lB = 2147483647: r = iA ^ lB: out = out & "bin.Integer1.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: lB = -2147483648: r = iA + lB: out = out & "bin.Integer2.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: lB = -2147483648: r = iA - lB: out = out & "bin.Integer2.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: lB = -2147483648: r = iA * lB: out = out & "bin.Integer2.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: lB = -2147483648: r = iA / lB: out = out & "bin.Integer2.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: lB = -2147483648: r = iA \ lB: out = out & "bin.Integer2.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: lB = -2147483648: r = iA Mod lB: out = out & "bin.Integer2.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: lB = -2147483648: r = iA ^ lB: out = out & "bin.Integer2.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: lB = 65536: r = iA + lB: out = out & "bin.Integer2.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: lB = 65536: r = iA - lB: out = out & "bin.Integer2.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: lB = 65536: r = iA * lB: out = out & "bin.Integer2.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: lB = 65536: r = iA / lB: out = out & "bin.Integer2.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: lB = 65536: r = iA \ lB: out = out & "bin.Integer2.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: lB = 65536: r = iA Mod lB: out = out & "bin.Integer2.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: lB = 65536: r = iA ^ lB: out = out & "bin.Integer2.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: lB = 2147483647: r = iA + lB: out = out & "bin.Integer2.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: lB = 2147483647: r = iA - lB: out = out & "bin.Integer2.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: lB = 2147483647: r = iA * lB: out = out & "bin.Integer2.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: lB = 2147483647: r = iA / lB: out = out & "bin.Integer2.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: lB = 2147483647: r = iA \ lB: out = out & "bin.Integer2.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: lB = 2147483647: r = iA Mod lB: out = out & "bin.Integer2.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: lB = 2147483647: r = iA ^ lB: out = out & "bin.Integer2.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: qB = -9223372036854775807^: r = iA + qB: out = out & "bin.Integer0.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: qB = -9223372036854775807^: r = iA - qB: out = out & "bin.Integer0.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: qB = -9223372036854775807^: r = iA * qB: out = out & "bin.Integer0.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: qB = -9223372036854775807^: r = iA / qB: out = out & "bin.Integer0.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: qB = -9223372036854775807^: r = iA \ qB: out = out & "bin.Integer0.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: qB = -9223372036854775807^: r = iA Mod qB: out = out & "bin.Integer0.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: qB = -9223372036854775807^: r = iA ^ qB: out = out & "bin.Integer0.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: qB = 3: r = iA + qB: out = out & "bin.Integer0.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: qB = 3: r = iA - qB: out = out & "bin.Integer0.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: qB = 3: r = iA * qB: out = out & "bin.Integer0.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: qB = 3: r = iA / qB: out = out & "bin.Integer0.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: qB = 3: r = iA \ qB: out = out & "bin.Integer0.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: qB = 3: r = iA Mod qB: out = out & "bin.Integer0.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: qB = 3: r = iA ^ qB: out = out & "bin.Integer0.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: qB = 9223372036854775807^: r = iA + qB: out = out & "bin.Integer0.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: qB = 9223372036854775807^: r = iA - qB: out = out & "bin.Integer0.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: qB = 9223372036854775807^: r = iA * qB: out = out & "bin.Integer0.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: qB = 9223372036854775807^: r = iA / qB: out = out & "bin.Integer0.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: qB = 9223372036854775807^: r = iA \ qB: out = out & "bin.Integer0.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: qB = 9223372036854775807^: r = iA Mod qB: out = out & "bin.Integer0.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: qB = 9223372036854775807^: r = iA ^ qB: out = out & "bin.Integer0.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: qB = -9223372036854775807^: r = iA + qB: out = out & "bin.Integer1.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: qB = -9223372036854775807^: r = iA - qB: out = out & "bin.Integer1.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: qB = -9223372036854775807^: r = iA * qB: out = out & "bin.Integer1.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: qB = -9223372036854775807^: r = iA / qB: out = out & "bin.Integer1.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: qB = -9223372036854775807^: r = iA \ qB: out = out & "bin.Integer1.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: qB = -9223372036854775807^: r = iA Mod qB: out = out & "bin.Integer1.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  P005 = out
End Function
Function P006() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: iA = 1: qB = -9223372036854775807^: r = iA ^ qB: out = out & "bin.Integer1.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: qB = 3: r = iA + qB: out = out & "bin.Integer1.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: qB = 3: r = iA - qB: out = out & "bin.Integer1.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: qB = 3: r = iA * qB: out = out & "bin.Integer1.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: qB = 3: r = iA / qB: out = out & "bin.Integer1.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: qB = 3: r = iA \ qB: out = out & "bin.Integer1.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: qB = 3: r = iA Mod qB: out = out & "bin.Integer1.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: qB = 3: r = iA ^ qB: out = out & "bin.Integer1.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: qB = 9223372036854775807^: r = iA + qB: out = out & "bin.Integer1.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: qB = 9223372036854775807^: r = iA - qB: out = out & "bin.Integer1.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: qB = 9223372036854775807^: r = iA * qB: out = out & "bin.Integer1.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: qB = 9223372036854775807^: r = iA / qB: out = out & "bin.Integer1.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: qB = 9223372036854775807^: r = iA \ qB: out = out & "bin.Integer1.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: qB = 9223372036854775807^: r = iA Mod qB: out = out & "bin.Integer1.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: qB = 9223372036854775807^: r = iA ^ qB: out = out & "bin.Integer1.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: qB = -9223372036854775807^: r = iA + qB: out = out & "bin.Integer2.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: qB = -9223372036854775807^: r = iA - qB: out = out & "bin.Integer2.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: qB = -9223372036854775807^: r = iA * qB: out = out & "bin.Integer2.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: qB = -9223372036854775807^: r = iA / qB: out = out & "bin.Integer2.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: qB = -9223372036854775807^: r = iA \ qB: out = out & "bin.Integer2.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: qB = -9223372036854775807^: r = iA Mod qB: out = out & "bin.Integer2.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: qB = -9223372036854775807^: r = iA ^ qB: out = out & "bin.Integer2.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: qB = 3: r = iA + qB: out = out & "bin.Integer2.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: qB = 3: r = iA - qB: out = out & "bin.Integer2.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: qB = 3: r = iA * qB: out = out & "bin.Integer2.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: qB = 3: r = iA / qB: out = out & "bin.Integer2.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: qB = 3: r = iA \ qB: out = out & "bin.Integer2.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: qB = 3: r = iA Mod qB: out = out & "bin.Integer2.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: qB = 3: r = iA ^ qB: out = out & "bin.Integer2.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: qB = 9223372036854775807^: r = iA + qB: out = out & "bin.Integer2.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: qB = 9223372036854775807^: r = iA - qB: out = out & "bin.Integer2.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: qB = 9223372036854775807^: r = iA * qB: out = out & "bin.Integer2.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: qB = 9223372036854775807^: r = iA / qB: out = out & "bin.Integer2.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: qB = 9223372036854775807^: r = iA \ qB: out = out & "bin.Integer2.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: qB = 9223372036854775807^: r = iA Mod qB: out = out & "bin.Integer2.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: qB = 9223372036854775807^: r = iA ^ qB: out = out & "bin.Integer2.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: sB = 0.5: r = iA + sB: out = out & "bin.Integer0.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: sB = 0.5: r = iA - sB: out = out & "bin.Integer0.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: sB = 0.5: r = iA * sB: out = out & "bin.Integer0.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: sB = 0.5: r = iA / sB: out = out & "bin.Integer0.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: sB = 0.5: r = iA \ sB: out = out & "bin.Integer0.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: sB = 0.5: r = iA Mod sB: out = out & "bin.Integer0.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: sB = 0.5: r = iA ^ sB: out = out & "bin.Integer0.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: sB = 16777216: r = iA + sB: out = out & "bin.Integer0.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: sB = 16777216: r = iA - sB: out = out & "bin.Integer0.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: sB = 16777216: r = iA * sB: out = out & "bin.Integer0.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: sB = 16777216: r = iA / sB: out = out & "bin.Integer0.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: sB = 16777216: r = iA \ sB: out = out & "bin.Integer0.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: sB = 16777216: r = iA Mod sB: out = out & "bin.Integer0.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: sB = 16777216: r = iA ^ sB: out = out & "bin.Integer0.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: sB = 3.0E+38: r = iA + sB: out = out & "bin.Integer0.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: sB = 3.0E+38: r = iA - sB: out = out & "bin.Integer0.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: sB = 3.0E+38: r = iA * sB: out = out & "bin.Integer0.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: sB = 3.0E+38: r = iA / sB: out = out & "bin.Integer0.div.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: sB = 3.0E+38: r = iA \ sB: out = out & "bin.Integer0.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: sB = 3.0E+38: r = iA Mod sB: out = out & "bin.Integer0.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: sB = 3.0E+38: r = iA ^ sB: out = out & "bin.Integer0.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: sB = 0.5: r = iA + sB: out = out & "bin.Integer1.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: sB = 0.5: r = iA - sB: out = out & "bin.Integer1.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: sB = 0.5: r = iA * sB: out = out & "bin.Integer1.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: sB = 0.5: r = iA / sB: out = out & "bin.Integer1.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: sB = 0.5: r = iA \ sB: out = out & "bin.Integer1.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: sB = 0.5: r = iA Mod sB: out = out & "bin.Integer1.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: sB = 0.5: r = iA ^ sB: out = out & "bin.Integer1.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: sB = 16777216: r = iA + sB: out = out & "bin.Integer1.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: sB = 16777216: r = iA - sB: out = out & "bin.Integer1.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: sB = 16777216: r = iA * sB: out = out & "bin.Integer1.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: sB = 16777216: r = iA / sB: out = out & "bin.Integer1.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: sB = 16777216: r = iA \ sB: out = out & "bin.Integer1.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: sB = 16777216: r = iA Mod sB: out = out & "bin.Integer1.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: sB = 16777216: r = iA ^ sB: out = out & "bin.Integer1.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: sB = 3.0E+38: r = iA + sB: out = out & "bin.Integer1.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: sB = 3.0E+38: r = iA - sB: out = out & "bin.Integer1.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: sB = 3.0E+38: r = iA * sB: out = out & "bin.Integer1.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: sB = 3.0E+38: r = iA / sB: out = out & "bin.Integer1.div.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: sB = 3.0E+38: r = iA \ sB: out = out & "bin.Integer1.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: sB = 3.0E+38: r = iA Mod sB: out = out & "bin.Integer1.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: sB = 3.0E+38: r = iA ^ sB: out = out & "bin.Integer1.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: sB = 0.5: r = iA + sB: out = out & "bin.Integer2.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: sB = 0.5: r = iA - sB: out = out & "bin.Integer2.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: sB = 0.5: r = iA * sB: out = out & "bin.Integer2.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: sB = 0.5: r = iA / sB: out = out & "bin.Integer2.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: sB = 0.5: r = iA \ sB: out = out & "bin.Integer2.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: sB = 0.5: r = iA Mod sB: out = out & "bin.Integer2.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: sB = 0.5: r = iA ^ sB: out = out & "bin.Integer2.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: sB = 16777216: r = iA + sB: out = out & "bin.Integer2.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: sB = 16777216: r = iA - sB: out = out & "bin.Integer2.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: sB = 16777216: r = iA * sB: out = out & "bin.Integer2.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: sB = 16777216: r = iA / sB: out = out & "bin.Integer2.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: sB = 16777216: r = iA \ sB: out = out & "bin.Integer2.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: sB = 16777216: r = iA Mod sB: out = out & "bin.Integer2.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: sB = 16777216: r = iA ^ sB: out = out & "bin.Integer2.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: sB = 3.0E+38: r = iA + sB: out = out & "bin.Integer2.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: sB = 3.0E+38: r = iA - sB: out = out & "bin.Integer2.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: sB = 3.0E+38: r = iA * sB: out = out & "bin.Integer2.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: sB = 3.0E+38: r = iA / sB: out = out & "bin.Integer2.div.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: sB = 3.0E+38: r = iA \ sB: out = out & "bin.Integer2.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: sB = 3.0E+38: r = iA Mod sB: out = out & "bin.Integer2.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: sB = 3.0E+38: r = iA ^ sB: out = out & "bin.Integer2.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: dB = 0.5: r = iA + dB: out = out & "bin.Integer0.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: dB = 0.5: r = iA - dB: out = out & "bin.Integer0.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: dB = 0.5: r = iA * dB: out = out & "bin.Integer0.mul.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: dB = 0.5: r = iA / dB: out = out & "bin.Integer0.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: dB = 0.5: r = iA \ dB: out = out & "bin.Integer0.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: dB = 0.5: r = iA Mod dB: out = out & "bin.Integer0.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: dB = 0.5: r = iA ^ dB: out = out & "bin.Integer0.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: dB = 2.5: r = iA + dB: out = out & "bin.Integer0.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: dB = 2.5: r = iA - dB: out = out & "bin.Integer0.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: dB = 2.5: r = iA * dB: out = out & "bin.Integer0.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: dB = 2.5: r = iA / dB: out = out & "bin.Integer0.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: dB = 2.5: r = iA \ dB: out = out & "bin.Integer0.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: dB = 2.5: r = iA Mod dB: out = out & "bin.Integer0.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: dB = 2.5: r = iA ^ dB: out = out & "bin.Integer0.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: dB = 1.0E+308: r = iA + dB: out = out & "bin.Integer0.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: dB = 1.0E+308: r = iA - dB: out = out & "bin.Integer0.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: dB = 1.0E+308: r = iA * dB: out = out & "bin.Integer0.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: dB = 1.0E+308: r = iA / dB: out = out & "bin.Integer0.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: dB = 1.0E+308: r = iA \ dB: out = out & "bin.Integer0.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: dB = 1.0E+308: r = iA Mod dB: out = out & "bin.Integer0.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: dB = 1.0E+308: r = iA ^ dB: out = out & "bin.Integer0.pow.Double2=" & Rz(Err.Number, r) & vbLf
  P006 = out
End Function
Function P007() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: iA = 1: dB = 0.5: r = iA + dB: out = out & "bin.Integer1.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: dB = 0.5: r = iA - dB: out = out & "bin.Integer1.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: dB = 0.5: r = iA * dB: out = out & "bin.Integer1.mul.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: dB = 0.5: r = iA / dB: out = out & "bin.Integer1.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: dB = 0.5: r = iA \ dB: out = out & "bin.Integer1.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: dB = 0.5: r = iA Mod dB: out = out & "bin.Integer1.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: dB = 0.5: r = iA ^ dB: out = out & "bin.Integer1.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: dB = 2.5: r = iA + dB: out = out & "bin.Integer1.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: dB = 2.5: r = iA - dB: out = out & "bin.Integer1.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: dB = 2.5: r = iA * dB: out = out & "bin.Integer1.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: dB = 2.5: r = iA / dB: out = out & "bin.Integer1.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: dB = 2.5: r = iA \ dB: out = out & "bin.Integer1.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: dB = 2.5: r = iA Mod dB: out = out & "bin.Integer1.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: dB = 2.5: r = iA ^ dB: out = out & "bin.Integer1.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: dB = 1.0E+308: r = iA + dB: out = out & "bin.Integer1.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: dB = 1.0E+308: r = iA - dB: out = out & "bin.Integer1.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: dB = 1.0E+308: r = iA * dB: out = out & "bin.Integer1.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: dB = 1.0E+308: r = iA / dB: out = out & "bin.Integer1.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: dB = 1.0E+308: r = iA \ dB: out = out & "bin.Integer1.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: dB = 1.0E+308: r = iA Mod dB: out = out & "bin.Integer1.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: dB = 1.0E+308: r = iA ^ dB: out = out & "bin.Integer1.pow.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: dB = 0.5: r = iA + dB: out = out & "bin.Integer2.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: dB = 0.5: r = iA - dB: out = out & "bin.Integer2.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: dB = 0.5: r = iA * dB: out = out & "bin.Integer2.mul.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: dB = 0.5: r = iA / dB: out = out & "bin.Integer2.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: dB = 0.5: r = iA \ dB: out = out & "bin.Integer2.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: dB = 0.5: r = iA Mod dB: out = out & "bin.Integer2.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: dB = 0.5: r = iA ^ dB: out = out & "bin.Integer2.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: dB = 2.5: r = iA + dB: out = out & "bin.Integer2.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: dB = 2.5: r = iA - dB: out = out & "bin.Integer2.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: dB = 2.5: r = iA * dB: out = out & "bin.Integer2.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: dB = 2.5: r = iA / dB: out = out & "bin.Integer2.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: dB = 2.5: r = iA \ dB: out = out & "bin.Integer2.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: dB = 2.5: r = iA Mod dB: out = out & "bin.Integer2.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: dB = 2.5: r = iA ^ dB: out = out & "bin.Integer2.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: dB = 1.0E+308: r = iA + dB: out = out & "bin.Integer2.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: dB = 1.0E+308: r = iA - dB: out = out & "bin.Integer2.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: dB = 1.0E+308: r = iA * dB: out = out & "bin.Integer2.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: dB = 1.0E+308: r = iA / dB: out = out & "bin.Integer2.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: dB = 1.0E+308: r = iA \ dB: out = out & "bin.Integer2.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: dB = 1.0E+308: r = iA Mod dB: out = out & "bin.Integer2.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: dB = 1.0E+308: r = iA ^ dB: out = out & "bin.Integer2.pow.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: cB = -0.5@: r = iA + cB: out = out & "bin.Integer0.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: cB = -0.5@: r = iA - cB: out = out & "bin.Integer0.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: cB = -0.5@: r = iA * cB: out = out & "bin.Integer0.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: cB = -0.5@: r = iA / cB: out = out & "bin.Integer0.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: cB = -0.5@: r = iA \ cB: out = out & "bin.Integer0.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: cB = -0.5@: r = iA Mod cB: out = out & "bin.Integer0.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: cB = -0.5@: r = iA ^ cB: out = out & "bin.Integer0.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: cB = 1.5@: r = iA + cB: out = out & "bin.Integer0.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: cB = 1.5@: r = iA - cB: out = out & "bin.Integer0.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: cB = 1.5@: r = iA * cB: out = out & "bin.Integer0.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: cB = 1.5@: r = iA / cB: out = out & "bin.Integer0.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: cB = 1.5@: r = iA \ cB: out = out & "bin.Integer0.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: cB = 1.5@: r = iA Mod cB: out = out & "bin.Integer0.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: cB = 1.5@: r = iA ^ cB: out = out & "bin.Integer0.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: cB = 922337203685477.5807@: r = iA + cB: out = out & "bin.Integer0.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: cB = 922337203685477.5807@: r = iA - cB: out = out & "bin.Integer0.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: cB = 922337203685477.5807@: r = iA * cB: out = out & "bin.Integer0.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: cB = 922337203685477.5807@: r = iA / cB: out = out & "bin.Integer0.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: cB = 922337203685477.5807@: r = iA \ cB: out = out & "bin.Integer0.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: cB = 922337203685477.5807@: r = iA Mod cB: out = out & "bin.Integer0.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: cB = 922337203685477.5807@: r = iA ^ cB: out = out & "bin.Integer0.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: cB = -0.5@: r = iA + cB: out = out & "bin.Integer1.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: cB = -0.5@: r = iA - cB: out = out & "bin.Integer1.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: cB = -0.5@: r = iA * cB: out = out & "bin.Integer1.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: cB = -0.5@: r = iA / cB: out = out & "bin.Integer1.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: cB = -0.5@: r = iA \ cB: out = out & "bin.Integer1.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: cB = -0.5@: r = iA Mod cB: out = out & "bin.Integer1.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: cB = -0.5@: r = iA ^ cB: out = out & "bin.Integer1.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: cB = 1.5@: r = iA + cB: out = out & "bin.Integer1.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: cB = 1.5@: r = iA - cB: out = out & "bin.Integer1.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: cB = 1.5@: r = iA * cB: out = out & "bin.Integer1.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: cB = 1.5@: r = iA / cB: out = out & "bin.Integer1.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: cB = 1.5@: r = iA \ cB: out = out & "bin.Integer1.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: cB = 1.5@: r = iA Mod cB: out = out & "bin.Integer1.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: cB = 1.5@: r = iA ^ cB: out = out & "bin.Integer1.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: cB = 922337203685477.5807@: r = iA + cB: out = out & "bin.Integer1.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: cB = 922337203685477.5807@: r = iA - cB: out = out & "bin.Integer1.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: cB = 922337203685477.5807@: r = iA * cB: out = out & "bin.Integer1.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: cB = 922337203685477.5807@: r = iA / cB: out = out & "bin.Integer1.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: cB = 922337203685477.5807@: r = iA \ cB: out = out & "bin.Integer1.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: cB = 922337203685477.5807@: r = iA Mod cB: out = out & "bin.Integer1.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: cB = 922337203685477.5807@: r = iA ^ cB: out = out & "bin.Integer1.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: cB = -0.5@: r = iA + cB: out = out & "bin.Integer2.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: cB = -0.5@: r = iA - cB: out = out & "bin.Integer2.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: cB = -0.5@: r = iA * cB: out = out & "bin.Integer2.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: cB = -0.5@: r = iA / cB: out = out & "bin.Integer2.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: cB = -0.5@: r = iA \ cB: out = out & "bin.Integer2.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: cB = -0.5@: r = iA Mod cB: out = out & "bin.Integer2.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: cB = -0.5@: r = iA ^ cB: out = out & "bin.Integer2.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: cB = 1.5@: r = iA + cB: out = out & "bin.Integer2.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: cB = 1.5@: r = iA - cB: out = out & "bin.Integer2.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: cB = 1.5@: r = iA * cB: out = out & "bin.Integer2.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: cB = 1.5@: r = iA / cB: out = out & "bin.Integer2.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: cB = 1.5@: r = iA \ cB: out = out & "bin.Integer2.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: cB = 1.5@: r = iA Mod cB: out = out & "bin.Integer2.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: cB = 1.5@: r = iA ^ cB: out = out & "bin.Integer2.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: cB = 922337203685477.5807@: r = iA + cB: out = out & "bin.Integer2.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: cB = 922337203685477.5807@: r = iA - cB: out = out & "bin.Integer2.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: cB = 922337203685477.5807@: r = iA * cB: out = out & "bin.Integer2.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: cB = 922337203685477.5807@: r = iA / cB: out = out & "bin.Integer2.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: cB = 922337203685477.5807@: r = iA \ cB: out = out & "bin.Integer2.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: cB = 922337203685477.5807@: r = iA Mod cB: out = out & "bin.Integer2.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: cB = 922337203685477.5807@: r = iA ^ cB: out = out & "bin.Integer2.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: vB = Empty: r = iA + vB: out = out & "bin.Integer0.add.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: vB = Empty: r = iA - vB: out = out & "bin.Integer0.sub.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: vB = Empty: r = iA * vB: out = out & "bin.Integer0.mul.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: vB = Empty: r = iA / vB: out = out & "bin.Integer0.div.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: vB = Empty: r = iA \ vB: out = out & "bin.Integer0.idiv.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: vB = Empty: r = iA Mod vB: out = out & "bin.Integer0.mod.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: vB = Empty: r = iA ^ vB: out = out & "bin.Integer0.pow.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: vB = 7: r = iA + vB: out = out & "bin.Integer0.add.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: vB = 7: r = iA - vB: out = out & "bin.Integer0.sub.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: vB = 7: r = iA * vB: out = out & "bin.Integer0.mul.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: vB = 7: r = iA / vB: out = out & "bin.Integer0.div.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: vB = 7: r = iA \ vB: out = out & "bin.Integer0.idiv.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: vB = 7: r = iA Mod vB: out = out & "bin.Integer0.mod.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: vB = 7: r = iA ^ vB: out = out & "bin.Integer0.pow.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: vB = 2147483647: r = iA + vB: out = out & "bin.Integer0.add.Variant2=" & Rz(Err.Number, r) & vbLf
  P007 = out
End Function
Function P008() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: iA = -32768: vB = 2147483647: r = iA - vB: out = out & "bin.Integer0.sub.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: vB = 2147483647: r = iA * vB: out = out & "bin.Integer0.mul.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: vB = 2147483647: r = iA / vB: out = out & "bin.Integer0.div.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: vB = 2147483647: r = iA \ vB: out = out & "bin.Integer0.idiv.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: vB = 2147483647: r = iA Mod vB: out = out & "bin.Integer0.mod.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: vB = 2147483647: r = iA ^ vB: out = out & "bin.Integer0.pow.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: vB = Empty: r = iA + vB: out = out & "bin.Integer1.add.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: vB = Empty: r = iA - vB: out = out & "bin.Integer1.sub.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: vB = Empty: r = iA * vB: out = out & "bin.Integer1.mul.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: vB = Empty: r = iA / vB: out = out & "bin.Integer1.div.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: vB = Empty: r = iA \ vB: out = out & "bin.Integer1.idiv.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: vB = Empty: r = iA Mod vB: out = out & "bin.Integer1.mod.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: vB = Empty: r = iA ^ vB: out = out & "bin.Integer1.pow.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: vB = 7: r = iA + vB: out = out & "bin.Integer1.add.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: vB = 7: r = iA - vB: out = out & "bin.Integer1.sub.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: vB = 7: r = iA * vB: out = out & "bin.Integer1.mul.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: vB = 7: r = iA / vB: out = out & "bin.Integer1.div.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: vB = 7: r = iA \ vB: out = out & "bin.Integer1.idiv.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: vB = 7: r = iA Mod vB: out = out & "bin.Integer1.mod.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: vB = 7: r = iA ^ vB: out = out & "bin.Integer1.pow.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: vB = 2147483647: r = iA + vB: out = out & "bin.Integer1.add.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: vB = 2147483647: r = iA - vB: out = out & "bin.Integer1.sub.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: vB = 2147483647: r = iA * vB: out = out & "bin.Integer1.mul.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: vB = 2147483647: r = iA / vB: out = out & "bin.Integer1.div.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: vB = 2147483647: r = iA \ vB: out = out & "bin.Integer1.idiv.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: vB = 2147483647: r = iA Mod vB: out = out & "bin.Integer1.mod.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: vB = 2147483647: r = iA ^ vB: out = out & "bin.Integer1.pow.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: vB = Empty: r = iA + vB: out = out & "bin.Integer2.add.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: vB = Empty: r = iA - vB: out = out & "bin.Integer2.sub.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: vB = Empty: r = iA * vB: out = out & "bin.Integer2.mul.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: vB = Empty: r = iA / vB: out = out & "bin.Integer2.div.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: vB = Empty: r = iA \ vB: out = out & "bin.Integer2.idiv.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: vB = Empty: r = iA Mod vB: out = out & "bin.Integer2.mod.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: vB = Empty: r = iA ^ vB: out = out & "bin.Integer2.pow.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: vB = 7: r = iA + vB: out = out & "bin.Integer2.add.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: vB = 7: r = iA - vB: out = out & "bin.Integer2.sub.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: vB = 7: r = iA * vB: out = out & "bin.Integer2.mul.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: vB = 7: r = iA / vB: out = out & "bin.Integer2.div.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: vB = 7: r = iA \ vB: out = out & "bin.Integer2.idiv.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: vB = 7: r = iA Mod vB: out = out & "bin.Integer2.mod.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: vB = 7: r = iA ^ vB: out = out & "bin.Integer2.pow.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: vB = 2147483647: r = iA + vB: out = out & "bin.Integer2.add.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: vB = 2147483647: r = iA - vB: out = out & "bin.Integer2.sub.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: vB = 2147483647: r = iA * vB: out = out & "bin.Integer2.mul.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: vB = 2147483647: r = iA / vB: out = out & "bin.Integer2.div.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: vB = 2147483647: r = iA \ vB: out = out & "bin.Integer2.idiv.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: vB = 2147483647: r = iA Mod vB: out = out & "bin.Integer2.mod.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: vB = 2147483647: r = iA ^ vB: out = out & "bin.Integer2.pow.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: bB = 0: r = lA + bB: out = out & "bin.Long0.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: bB = 0: r = lA - bB: out = out & "bin.Long0.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: bB = 0: r = lA * bB: out = out & "bin.Long0.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: bB = 0: r = lA / bB: out = out & "bin.Long0.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: bB = 0: r = lA \ bB: out = out & "bin.Long0.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: bB = 0: r = lA Mod bB: out = out & "bin.Long0.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: bB = 0: r = lA ^ bB: out = out & "bin.Long0.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: bB = 200: r = lA + bB: out = out & "bin.Long0.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: bB = 200: r = lA - bB: out = out & "bin.Long0.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: bB = 200: r = lA * bB: out = out & "bin.Long0.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: bB = 200: r = lA / bB: out = out & "bin.Long0.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: bB = 200: r = lA \ bB: out = out & "bin.Long0.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: bB = 200: r = lA Mod bB: out = out & "bin.Long0.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: bB = 200: r = lA ^ bB: out = out & "bin.Long0.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: bB = 255: r = lA + bB: out = out & "bin.Long0.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: bB = 255: r = lA - bB: out = out & "bin.Long0.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: bB = 255: r = lA * bB: out = out & "bin.Long0.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: bB = 255: r = lA / bB: out = out & "bin.Long0.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: bB = 255: r = lA \ bB: out = out & "bin.Long0.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: bB = 255: r = lA Mod bB: out = out & "bin.Long0.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: bB = 255: r = lA ^ bB: out = out & "bin.Long0.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: bB = 0: r = lA + bB: out = out & "bin.Long1.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: bB = 0: r = lA - bB: out = out & "bin.Long1.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: bB = 0: r = lA * bB: out = out & "bin.Long1.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: bB = 0: r = lA / bB: out = out & "bin.Long1.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: bB = 0: r = lA \ bB: out = out & "bin.Long1.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: bB = 0: r = lA Mod bB: out = out & "bin.Long1.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: bB = 0: r = lA ^ bB: out = out & "bin.Long1.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: bB = 200: r = lA + bB: out = out & "bin.Long1.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: bB = 200: r = lA - bB: out = out & "bin.Long1.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: bB = 200: r = lA * bB: out = out & "bin.Long1.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: bB = 200: r = lA / bB: out = out & "bin.Long1.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: bB = 200: r = lA \ bB: out = out & "bin.Long1.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: bB = 200: r = lA Mod bB: out = out & "bin.Long1.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: bB = 200: r = lA ^ bB: out = out & "bin.Long1.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: bB = 255: r = lA + bB: out = out & "bin.Long1.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: bB = 255: r = lA - bB: out = out & "bin.Long1.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: bB = 255: r = lA * bB: out = out & "bin.Long1.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: bB = 255: r = lA / bB: out = out & "bin.Long1.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: bB = 255: r = lA \ bB: out = out & "bin.Long1.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: bB = 255: r = lA Mod bB: out = out & "bin.Long1.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: bB = 255: r = lA ^ bB: out = out & "bin.Long1.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: bB = 0: r = lA + bB: out = out & "bin.Long2.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: bB = 0: r = lA - bB: out = out & "bin.Long2.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: bB = 0: r = lA * bB: out = out & "bin.Long2.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: bB = 0: r = lA / bB: out = out & "bin.Long2.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: bB = 0: r = lA \ bB: out = out & "bin.Long2.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: bB = 0: r = lA Mod bB: out = out & "bin.Long2.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: bB = 0: r = lA ^ bB: out = out & "bin.Long2.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: bB = 200: r = lA + bB: out = out & "bin.Long2.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: bB = 200: r = lA - bB: out = out & "bin.Long2.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: bB = 200: r = lA * bB: out = out & "bin.Long2.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: bB = 200: r = lA / bB: out = out & "bin.Long2.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: bB = 200: r = lA \ bB: out = out & "bin.Long2.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: bB = 200: r = lA Mod bB: out = out & "bin.Long2.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: bB = 200: r = lA ^ bB: out = out & "bin.Long2.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: bB = 255: r = lA + bB: out = out & "bin.Long2.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: bB = 255: r = lA - bB: out = out & "bin.Long2.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: bB = 255: r = lA * bB: out = out & "bin.Long2.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: bB = 255: r = lA / bB: out = out & "bin.Long2.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: bB = 255: r = lA \ bB: out = out & "bin.Long2.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: bB = 255: r = lA Mod bB: out = out & "bin.Long2.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: bB = 255: r = lA ^ bB: out = out & "bin.Long2.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: iB = -32768: r = lA + iB: out = out & "bin.Long0.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: iB = -32768: r = lA - iB: out = out & "bin.Long0.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: iB = -32768: r = lA * iB: out = out & "bin.Long0.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: iB = -32768: r = lA / iB: out = out & "bin.Long0.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: iB = -32768: r = lA \ iB: out = out & "bin.Long0.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: iB = -32768: r = lA Mod iB: out = out & "bin.Long0.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: iB = -32768: r = lA ^ iB: out = out & "bin.Long0.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: iB = 1: r = lA + iB: out = out & "bin.Long0.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: iB = 1: r = lA - iB: out = out & "bin.Long0.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  P008 = out
End Function
Function P009() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: lA = -2147483648: iB = 1: r = lA * iB: out = out & "bin.Long0.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: iB = 1: r = lA / iB: out = out & "bin.Long0.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: iB = 1: r = lA \ iB: out = out & "bin.Long0.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: iB = 1: r = lA Mod iB: out = out & "bin.Long0.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: iB = 1: r = lA ^ iB: out = out & "bin.Long0.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: iB = 32767: r = lA + iB: out = out & "bin.Long0.add.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: iB = 32767: r = lA - iB: out = out & "bin.Long0.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: iB = 32767: r = lA * iB: out = out & "bin.Long0.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: iB = 32767: r = lA / iB: out = out & "bin.Long0.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: iB = 32767: r = lA \ iB: out = out & "bin.Long0.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: iB = 32767: r = lA Mod iB: out = out & "bin.Long0.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: iB = 32767: r = lA ^ iB: out = out & "bin.Long0.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: iB = -32768: r = lA + iB: out = out & "bin.Long1.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: iB = -32768: r = lA - iB: out = out & "bin.Long1.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: iB = -32768: r = lA * iB: out = out & "bin.Long1.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: iB = -32768: r = lA / iB: out = out & "bin.Long1.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: iB = -32768: r = lA \ iB: out = out & "bin.Long1.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: iB = -32768: r = lA Mod iB: out = out & "bin.Long1.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: iB = -32768: r = lA ^ iB: out = out & "bin.Long1.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: iB = 1: r = lA + iB: out = out & "bin.Long1.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: iB = 1: r = lA - iB: out = out & "bin.Long1.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: iB = 1: r = lA * iB: out = out & "bin.Long1.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: iB = 1: r = lA / iB: out = out & "bin.Long1.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: iB = 1: r = lA \ iB: out = out & "bin.Long1.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: iB = 1: r = lA Mod iB: out = out & "bin.Long1.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: iB = 1: r = lA ^ iB: out = out & "bin.Long1.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: iB = 32767: r = lA + iB: out = out & "bin.Long1.add.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: iB = 32767: r = lA - iB: out = out & "bin.Long1.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: iB = 32767: r = lA * iB: out = out & "bin.Long1.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: iB = 32767: r = lA / iB: out = out & "bin.Long1.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: iB = 32767: r = lA \ iB: out = out & "bin.Long1.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: iB = 32767: r = lA Mod iB: out = out & "bin.Long1.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: iB = 32767: r = lA ^ iB: out = out & "bin.Long1.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: iB = -32768: r = lA + iB: out = out & "bin.Long2.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: iB = -32768: r = lA - iB: out = out & "bin.Long2.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: iB = -32768: r = lA * iB: out = out & "bin.Long2.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: iB = -32768: r = lA / iB: out = out & "bin.Long2.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: iB = -32768: r = lA \ iB: out = out & "bin.Long2.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: iB = -32768: r = lA Mod iB: out = out & "bin.Long2.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: iB = -32768: r = lA ^ iB: out = out & "bin.Long2.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: iB = 1: r = lA + iB: out = out & "bin.Long2.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: iB = 1: r = lA - iB: out = out & "bin.Long2.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: iB = 1: r = lA * iB: out = out & "bin.Long2.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: iB = 1: r = lA / iB: out = out & "bin.Long2.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: iB = 1: r = lA \ iB: out = out & "bin.Long2.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: iB = 1: r = lA Mod iB: out = out & "bin.Long2.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: iB = 1: r = lA ^ iB: out = out & "bin.Long2.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: iB = 32767: r = lA + iB: out = out & "bin.Long2.add.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: iB = 32767: r = lA - iB: out = out & "bin.Long2.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: iB = 32767: r = lA * iB: out = out & "bin.Long2.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: iB = 32767: r = lA / iB: out = out & "bin.Long2.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: iB = 32767: r = lA \ iB: out = out & "bin.Long2.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: iB = 32767: r = lA Mod iB: out = out & "bin.Long2.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: iB = 32767: r = lA ^ iB: out = out & "bin.Long2.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: lB = -2147483648: r = lA + lB: out = out & "bin.Long0.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: lB = -2147483648: r = lA - lB: out = out & "bin.Long0.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: lB = -2147483648: r = lA * lB: out = out & "bin.Long0.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: lB = -2147483648: r = lA / lB: out = out & "bin.Long0.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: lB = -2147483648: r = lA \ lB: out = out & "bin.Long0.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: lB = -2147483648: r = lA Mod lB: out = out & "bin.Long0.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: lB = -2147483648: r = lA ^ lB: out = out & "bin.Long0.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: lB = 65536: r = lA + lB: out = out & "bin.Long0.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: lB = 65536: r = lA - lB: out = out & "bin.Long0.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: lB = 65536: r = lA * lB: out = out & "bin.Long0.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: lB = 65536: r = lA / lB: out = out & "bin.Long0.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: lB = 65536: r = lA \ lB: out = out & "bin.Long0.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: lB = 65536: r = lA Mod lB: out = out & "bin.Long0.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: lB = 65536: r = lA ^ lB: out = out & "bin.Long0.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: lB = 2147483647: r = lA + lB: out = out & "bin.Long0.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: lB = 2147483647: r = lA - lB: out = out & "bin.Long0.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: lB = 2147483647: r = lA * lB: out = out & "bin.Long0.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: lB = 2147483647: r = lA / lB: out = out & "bin.Long0.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: lB = 2147483647: r = lA \ lB: out = out & "bin.Long0.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: lB = 2147483647: r = lA Mod lB: out = out & "bin.Long0.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: lB = 2147483647: r = lA ^ lB: out = out & "bin.Long0.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: lB = -2147483648: r = lA + lB: out = out & "bin.Long1.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: lB = -2147483648: r = lA - lB: out = out & "bin.Long1.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: lB = -2147483648: r = lA * lB: out = out & "bin.Long1.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: lB = -2147483648: r = lA / lB: out = out & "bin.Long1.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: lB = -2147483648: r = lA \ lB: out = out & "bin.Long1.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: lB = -2147483648: r = lA Mod lB: out = out & "bin.Long1.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: lB = -2147483648: r = lA ^ lB: out = out & "bin.Long1.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: lB = 65536: r = lA + lB: out = out & "bin.Long1.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: lB = 65536: r = lA - lB: out = out & "bin.Long1.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: lB = 65536: r = lA * lB: out = out & "bin.Long1.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: lB = 65536: r = lA / lB: out = out & "bin.Long1.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: lB = 65536: r = lA \ lB: out = out & "bin.Long1.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: lB = 65536: r = lA Mod lB: out = out & "bin.Long1.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: lB = 65536: r = lA ^ lB: out = out & "bin.Long1.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: lB = 2147483647: r = lA + lB: out = out & "bin.Long1.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: lB = 2147483647: r = lA - lB: out = out & "bin.Long1.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: lB = 2147483647: r = lA * lB: out = out & "bin.Long1.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: lB = 2147483647: r = lA / lB: out = out & "bin.Long1.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: lB = 2147483647: r = lA \ lB: out = out & "bin.Long1.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: lB = 2147483647: r = lA Mod lB: out = out & "bin.Long1.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: lB = 2147483647: r = lA ^ lB: out = out & "bin.Long1.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: lB = -2147483648: r = lA + lB: out = out & "bin.Long2.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: lB = -2147483648: r = lA - lB: out = out & "bin.Long2.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: lB = -2147483648: r = lA * lB: out = out & "bin.Long2.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: lB = -2147483648: r = lA / lB: out = out & "bin.Long2.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: lB = -2147483648: r = lA \ lB: out = out & "bin.Long2.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: lB = -2147483648: r = lA Mod lB: out = out & "bin.Long2.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: lB = -2147483648: r = lA ^ lB: out = out & "bin.Long2.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: lB = 65536: r = lA + lB: out = out & "bin.Long2.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: lB = 65536: r = lA - lB: out = out & "bin.Long2.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: lB = 65536: r = lA * lB: out = out & "bin.Long2.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: lB = 65536: r = lA / lB: out = out & "bin.Long2.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: lB = 65536: r = lA \ lB: out = out & "bin.Long2.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: lB = 65536: r = lA Mod lB: out = out & "bin.Long2.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: lB = 65536: r = lA ^ lB: out = out & "bin.Long2.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: lB = 2147483647: r = lA + lB: out = out & "bin.Long2.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: lB = 2147483647: r = lA - lB: out = out & "bin.Long2.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: lB = 2147483647: r = lA * lB: out = out & "bin.Long2.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: lB = 2147483647: r = lA / lB: out = out & "bin.Long2.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: lB = 2147483647: r = lA \ lB: out = out & "bin.Long2.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: lB = 2147483647: r = lA Mod lB: out = out & "bin.Long2.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: lB = 2147483647: r = lA ^ lB: out = out & "bin.Long2.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: qB = -9223372036854775807^: r = lA + qB: out = out & "bin.Long0.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: qB = -9223372036854775807^: r = lA - qB: out = out & "bin.Long0.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: qB = -9223372036854775807^: r = lA * qB: out = out & "bin.Long0.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  P009 = out
End Function
Function P010() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: lA = -2147483648: qB = -9223372036854775807^: r = lA / qB: out = out & "bin.Long0.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: qB = -9223372036854775807^: r = lA \ qB: out = out & "bin.Long0.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: qB = -9223372036854775807^: r = lA Mod qB: out = out & "bin.Long0.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: qB = -9223372036854775807^: r = lA ^ qB: out = out & "bin.Long0.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: qB = 3: r = lA + qB: out = out & "bin.Long0.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: qB = 3: r = lA - qB: out = out & "bin.Long0.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: qB = 3: r = lA * qB: out = out & "bin.Long0.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: qB = 3: r = lA / qB: out = out & "bin.Long0.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: qB = 3: r = lA \ qB: out = out & "bin.Long0.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: qB = 3: r = lA Mod qB: out = out & "bin.Long0.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: qB = 3: r = lA ^ qB: out = out & "bin.Long0.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: qB = 9223372036854775807^: r = lA + qB: out = out & "bin.Long0.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: qB = 9223372036854775807^: r = lA - qB: out = out & "bin.Long0.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: qB = 9223372036854775807^: r = lA * qB: out = out & "bin.Long0.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: qB = 9223372036854775807^: r = lA / qB: out = out & "bin.Long0.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: qB = 9223372036854775807^: r = lA \ qB: out = out & "bin.Long0.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: qB = 9223372036854775807^: r = lA Mod qB: out = out & "bin.Long0.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: qB = 9223372036854775807^: r = lA ^ qB: out = out & "bin.Long0.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: qB = -9223372036854775807^: r = lA + qB: out = out & "bin.Long1.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: qB = -9223372036854775807^: r = lA - qB: out = out & "bin.Long1.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: qB = -9223372036854775807^: r = lA * qB: out = out & "bin.Long1.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: qB = -9223372036854775807^: r = lA / qB: out = out & "bin.Long1.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: qB = -9223372036854775807^: r = lA \ qB: out = out & "bin.Long1.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: qB = -9223372036854775807^: r = lA Mod qB: out = out & "bin.Long1.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: qB = -9223372036854775807^: r = lA ^ qB: out = out & "bin.Long1.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: qB = 3: r = lA + qB: out = out & "bin.Long1.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: qB = 3: r = lA - qB: out = out & "bin.Long1.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: qB = 3: r = lA * qB: out = out & "bin.Long1.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: qB = 3: r = lA / qB: out = out & "bin.Long1.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: qB = 3: r = lA \ qB: out = out & "bin.Long1.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: qB = 3: r = lA Mod qB: out = out & "bin.Long1.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: qB = 3: r = lA ^ qB: out = out & "bin.Long1.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: qB = 9223372036854775807^: r = lA + qB: out = out & "bin.Long1.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: qB = 9223372036854775807^: r = lA - qB: out = out & "bin.Long1.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: qB = 9223372036854775807^: r = lA * qB: out = out & "bin.Long1.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: qB = 9223372036854775807^: r = lA / qB: out = out & "bin.Long1.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: qB = 9223372036854775807^: r = lA \ qB: out = out & "bin.Long1.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: qB = 9223372036854775807^: r = lA Mod qB: out = out & "bin.Long1.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: qB = 9223372036854775807^: r = lA ^ qB: out = out & "bin.Long1.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: qB = -9223372036854775807^: r = lA + qB: out = out & "bin.Long2.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: qB = -9223372036854775807^: r = lA - qB: out = out & "bin.Long2.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: qB = -9223372036854775807^: r = lA * qB: out = out & "bin.Long2.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: qB = -9223372036854775807^: r = lA / qB: out = out & "bin.Long2.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: qB = -9223372036854775807^: r = lA \ qB: out = out & "bin.Long2.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: qB = -9223372036854775807^: r = lA Mod qB: out = out & "bin.Long2.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: qB = -9223372036854775807^: r = lA ^ qB: out = out & "bin.Long2.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: qB = 3: r = lA + qB: out = out & "bin.Long2.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: qB = 3: r = lA - qB: out = out & "bin.Long2.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: qB = 3: r = lA * qB: out = out & "bin.Long2.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: qB = 3: r = lA / qB: out = out & "bin.Long2.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: qB = 3: r = lA \ qB: out = out & "bin.Long2.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: qB = 3: r = lA Mod qB: out = out & "bin.Long2.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: qB = 3: r = lA ^ qB: out = out & "bin.Long2.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: qB = 9223372036854775807^: r = lA + qB: out = out & "bin.Long2.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: qB = 9223372036854775807^: r = lA - qB: out = out & "bin.Long2.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: qB = 9223372036854775807^: r = lA * qB: out = out & "bin.Long2.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: qB = 9223372036854775807^: r = lA / qB: out = out & "bin.Long2.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: qB = 9223372036854775807^: r = lA \ qB: out = out & "bin.Long2.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: qB = 9223372036854775807^: r = lA Mod qB: out = out & "bin.Long2.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: qB = 9223372036854775807^: r = lA ^ qB: out = out & "bin.Long2.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: sB = 0.5: r = lA + sB: out = out & "bin.Long0.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: sB = 0.5: r = lA - sB: out = out & "bin.Long0.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: sB = 0.5: r = lA * sB: out = out & "bin.Long0.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: sB = 0.5: r = lA / sB: out = out & "bin.Long0.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: sB = 0.5: r = lA \ sB: out = out & "bin.Long0.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: sB = 0.5: r = lA Mod sB: out = out & "bin.Long0.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: sB = 0.5: r = lA ^ sB: out = out & "bin.Long0.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: sB = 16777216: r = lA + sB: out = out & "bin.Long0.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: sB = 16777216: r = lA - sB: out = out & "bin.Long0.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: sB = 16777216: r = lA * sB: out = out & "bin.Long0.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: sB = 16777216: r = lA / sB: out = out & "bin.Long0.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: sB = 16777216: r = lA \ sB: out = out & "bin.Long0.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: sB = 16777216: r = lA Mod sB: out = out & "bin.Long0.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: sB = 16777216: r = lA ^ sB: out = out & "bin.Long0.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: sB = 3.0E+38: r = lA + sB: out = out & "bin.Long0.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: sB = 3.0E+38: r = lA - sB: out = out & "bin.Long0.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: sB = 3.0E+38: r = lA * sB: out = out & "bin.Long0.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: sB = 3.0E+38: r = lA / sB: out = out & "bin.Long0.div.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: sB = 3.0E+38: r = lA \ sB: out = out & "bin.Long0.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: sB = 3.0E+38: r = lA Mod sB: out = out & "bin.Long0.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: sB = 3.0E+38: r = lA ^ sB: out = out & "bin.Long0.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: sB = 0.5: r = lA + sB: out = out & "bin.Long1.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: sB = 0.5: r = lA - sB: out = out & "bin.Long1.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: sB = 0.5: r = lA * sB: out = out & "bin.Long1.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: sB = 0.5: r = lA / sB: out = out & "bin.Long1.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: sB = 0.5: r = lA \ sB: out = out & "bin.Long1.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: sB = 0.5: r = lA Mod sB: out = out & "bin.Long1.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: sB = 0.5: r = lA ^ sB: out = out & "bin.Long1.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: sB = 16777216: r = lA + sB: out = out & "bin.Long1.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: sB = 16777216: r = lA - sB: out = out & "bin.Long1.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: sB = 16777216: r = lA * sB: out = out & "bin.Long1.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: sB = 16777216: r = lA / sB: out = out & "bin.Long1.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: sB = 16777216: r = lA \ sB: out = out & "bin.Long1.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: sB = 16777216: r = lA Mod sB: out = out & "bin.Long1.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: sB = 16777216: r = lA ^ sB: out = out & "bin.Long1.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: sB = 3.0E+38: r = lA + sB: out = out & "bin.Long1.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: sB = 3.0E+38: r = lA - sB: out = out & "bin.Long1.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: sB = 3.0E+38: r = lA * sB: out = out & "bin.Long1.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: sB = 3.0E+38: r = lA / sB: out = out & "bin.Long1.div.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: sB = 3.0E+38: r = lA \ sB: out = out & "bin.Long1.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: sB = 3.0E+38: r = lA Mod sB: out = out & "bin.Long1.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: sB = 3.0E+38: r = lA ^ sB: out = out & "bin.Long1.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: sB = 0.5: r = lA + sB: out = out & "bin.Long2.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: sB = 0.5: r = lA - sB: out = out & "bin.Long2.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: sB = 0.5: r = lA * sB: out = out & "bin.Long2.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: sB = 0.5: r = lA / sB: out = out & "bin.Long2.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: sB = 0.5: r = lA \ sB: out = out & "bin.Long2.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: sB = 0.5: r = lA Mod sB: out = out & "bin.Long2.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: sB = 0.5: r = lA ^ sB: out = out & "bin.Long2.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: sB = 16777216: r = lA + sB: out = out & "bin.Long2.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: sB = 16777216: r = lA - sB: out = out & "bin.Long2.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: sB = 16777216: r = lA * sB: out = out & "bin.Long2.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: sB = 16777216: r = lA / sB: out = out & "bin.Long2.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: sB = 16777216: r = lA \ sB: out = out & "bin.Long2.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: sB = 16777216: r = lA Mod sB: out = out & "bin.Long2.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: sB = 16777216: r = lA ^ sB: out = out & "bin.Long2.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: sB = 3.0E+38: r = lA + sB: out = out & "bin.Long2.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: sB = 3.0E+38: r = lA - sB: out = out & "bin.Long2.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: sB = 3.0E+38: r = lA * sB: out = out & "bin.Long2.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: sB = 3.0E+38: r = lA / sB: out = out & "bin.Long2.div.Single2=" & Rz(Err.Number, r) & vbLf
  P010 = out
End Function
Function P011() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: lA = 2147483647: sB = 3.0E+38: r = lA \ sB: out = out & "bin.Long2.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: sB = 3.0E+38: r = lA Mod sB: out = out & "bin.Long2.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: sB = 3.0E+38: r = lA ^ sB: out = out & "bin.Long2.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: dB = 0.5: r = lA + dB: out = out & "bin.Long0.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: dB = 0.5: r = lA - dB: out = out & "bin.Long0.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: dB = 0.5: r = lA * dB: out = out & "bin.Long0.mul.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: dB = 0.5: r = lA / dB: out = out & "bin.Long0.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: dB = 0.5: r = lA \ dB: out = out & "bin.Long0.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: dB = 0.5: r = lA Mod dB: out = out & "bin.Long0.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: dB = 0.5: r = lA ^ dB: out = out & "bin.Long0.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: dB = 2.5: r = lA + dB: out = out & "bin.Long0.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: dB = 2.5: r = lA - dB: out = out & "bin.Long0.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: dB = 2.5: r = lA * dB: out = out & "bin.Long0.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: dB = 2.5: r = lA / dB: out = out & "bin.Long0.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: dB = 2.5: r = lA \ dB: out = out & "bin.Long0.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: dB = 2.5: r = lA Mod dB: out = out & "bin.Long0.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: dB = 2.5: r = lA ^ dB: out = out & "bin.Long0.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: dB = 1.0E+308: r = lA + dB: out = out & "bin.Long0.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: dB = 1.0E+308: r = lA - dB: out = out & "bin.Long0.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: dB = 1.0E+308: r = lA * dB: out = out & "bin.Long0.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: dB = 1.0E+308: r = lA / dB: out = out & "bin.Long0.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: dB = 1.0E+308: r = lA \ dB: out = out & "bin.Long0.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: dB = 1.0E+308: r = lA Mod dB: out = out & "bin.Long0.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: dB = 1.0E+308: r = lA ^ dB: out = out & "bin.Long0.pow.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: dB = 0.5: r = lA + dB: out = out & "bin.Long1.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: dB = 0.5: r = lA - dB: out = out & "bin.Long1.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: dB = 0.5: r = lA * dB: out = out & "bin.Long1.mul.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: dB = 0.5: r = lA / dB: out = out & "bin.Long1.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: dB = 0.5: r = lA \ dB: out = out & "bin.Long1.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: dB = 0.5: r = lA Mod dB: out = out & "bin.Long1.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: dB = 0.5: r = lA ^ dB: out = out & "bin.Long1.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: dB = 2.5: r = lA + dB: out = out & "bin.Long1.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: dB = 2.5: r = lA - dB: out = out & "bin.Long1.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: dB = 2.5: r = lA * dB: out = out & "bin.Long1.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: dB = 2.5: r = lA / dB: out = out & "bin.Long1.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: dB = 2.5: r = lA \ dB: out = out & "bin.Long1.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: dB = 2.5: r = lA Mod dB: out = out & "bin.Long1.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: dB = 2.5: r = lA ^ dB: out = out & "bin.Long1.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: dB = 1.0E+308: r = lA + dB: out = out & "bin.Long1.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: dB = 1.0E+308: r = lA - dB: out = out & "bin.Long1.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: dB = 1.0E+308: r = lA * dB: out = out & "bin.Long1.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: dB = 1.0E+308: r = lA / dB: out = out & "bin.Long1.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: dB = 1.0E+308: r = lA \ dB: out = out & "bin.Long1.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: dB = 1.0E+308: r = lA Mod dB: out = out & "bin.Long1.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: dB = 1.0E+308: r = lA ^ dB: out = out & "bin.Long1.pow.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: dB = 0.5: r = lA + dB: out = out & "bin.Long2.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: dB = 0.5: r = lA - dB: out = out & "bin.Long2.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: dB = 0.5: r = lA * dB: out = out & "bin.Long2.mul.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: dB = 0.5: r = lA / dB: out = out & "bin.Long2.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: dB = 0.5: r = lA \ dB: out = out & "bin.Long2.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: dB = 0.5: r = lA Mod dB: out = out & "bin.Long2.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: dB = 0.5: r = lA ^ dB: out = out & "bin.Long2.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: dB = 2.5: r = lA + dB: out = out & "bin.Long2.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: dB = 2.5: r = lA - dB: out = out & "bin.Long2.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: dB = 2.5: r = lA * dB: out = out & "bin.Long2.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: dB = 2.5: r = lA / dB: out = out & "bin.Long2.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: dB = 2.5: r = lA \ dB: out = out & "bin.Long2.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: dB = 2.5: r = lA Mod dB: out = out & "bin.Long2.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: dB = 2.5: r = lA ^ dB: out = out & "bin.Long2.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: dB = 1.0E+308: r = lA + dB: out = out & "bin.Long2.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: dB = 1.0E+308: r = lA - dB: out = out & "bin.Long2.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: dB = 1.0E+308: r = lA * dB: out = out & "bin.Long2.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: dB = 1.0E+308: r = lA / dB: out = out & "bin.Long2.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: dB = 1.0E+308: r = lA \ dB: out = out & "bin.Long2.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: dB = 1.0E+308: r = lA Mod dB: out = out & "bin.Long2.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: dB = 1.0E+308: r = lA ^ dB: out = out & "bin.Long2.pow.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: cB = -0.5@: r = lA + cB: out = out & "bin.Long0.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: cB = -0.5@: r = lA - cB: out = out & "bin.Long0.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: cB = -0.5@: r = lA * cB: out = out & "bin.Long0.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: cB = -0.5@: r = lA / cB: out = out & "bin.Long0.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: cB = -0.5@: r = lA \ cB: out = out & "bin.Long0.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: cB = -0.5@: r = lA Mod cB: out = out & "bin.Long0.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: cB = -0.5@: r = lA ^ cB: out = out & "bin.Long0.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: cB = 1.5@: r = lA + cB: out = out & "bin.Long0.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: cB = 1.5@: r = lA - cB: out = out & "bin.Long0.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: cB = 1.5@: r = lA * cB: out = out & "bin.Long0.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: cB = 1.5@: r = lA / cB: out = out & "bin.Long0.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: cB = 1.5@: r = lA \ cB: out = out & "bin.Long0.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: cB = 1.5@: r = lA Mod cB: out = out & "bin.Long0.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: cB = 1.5@: r = lA ^ cB: out = out & "bin.Long0.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: cB = 922337203685477.5807@: r = lA + cB: out = out & "bin.Long0.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: cB = 922337203685477.5807@: r = lA - cB: out = out & "bin.Long0.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: cB = 922337203685477.5807@: r = lA * cB: out = out & "bin.Long0.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: cB = 922337203685477.5807@: r = lA / cB: out = out & "bin.Long0.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: cB = 922337203685477.5807@: r = lA \ cB: out = out & "bin.Long0.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: cB = 922337203685477.5807@: r = lA Mod cB: out = out & "bin.Long0.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: cB = 922337203685477.5807@: r = lA ^ cB: out = out & "bin.Long0.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: cB = -0.5@: r = lA + cB: out = out & "bin.Long1.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: cB = -0.5@: r = lA - cB: out = out & "bin.Long1.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: cB = -0.5@: r = lA * cB: out = out & "bin.Long1.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: cB = -0.5@: r = lA / cB: out = out & "bin.Long1.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: cB = -0.5@: r = lA \ cB: out = out & "bin.Long1.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: cB = -0.5@: r = lA Mod cB: out = out & "bin.Long1.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: cB = -0.5@: r = lA ^ cB: out = out & "bin.Long1.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: cB = 1.5@: r = lA + cB: out = out & "bin.Long1.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: cB = 1.5@: r = lA - cB: out = out & "bin.Long1.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: cB = 1.5@: r = lA * cB: out = out & "bin.Long1.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: cB = 1.5@: r = lA / cB: out = out & "bin.Long1.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: cB = 1.5@: r = lA \ cB: out = out & "bin.Long1.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: cB = 1.5@: r = lA Mod cB: out = out & "bin.Long1.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: cB = 1.5@: r = lA ^ cB: out = out & "bin.Long1.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: cB = 922337203685477.5807@: r = lA + cB: out = out & "bin.Long1.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: cB = 922337203685477.5807@: r = lA - cB: out = out & "bin.Long1.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: cB = 922337203685477.5807@: r = lA * cB: out = out & "bin.Long1.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: cB = 922337203685477.5807@: r = lA / cB: out = out & "bin.Long1.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: cB = 922337203685477.5807@: r = lA \ cB: out = out & "bin.Long1.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: cB = 922337203685477.5807@: r = lA Mod cB: out = out & "bin.Long1.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: cB = 922337203685477.5807@: r = lA ^ cB: out = out & "bin.Long1.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: cB = -0.5@: r = lA + cB: out = out & "bin.Long2.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: cB = -0.5@: r = lA - cB: out = out & "bin.Long2.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: cB = -0.5@: r = lA * cB: out = out & "bin.Long2.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: cB = -0.5@: r = lA / cB: out = out & "bin.Long2.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: cB = -0.5@: r = lA \ cB: out = out & "bin.Long2.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: cB = -0.5@: r = lA Mod cB: out = out & "bin.Long2.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: cB = -0.5@: r = lA ^ cB: out = out & "bin.Long2.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: cB = 1.5@: r = lA + cB: out = out & "bin.Long2.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: cB = 1.5@: r = lA - cB: out = out & "bin.Long2.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: cB = 1.5@: r = lA * cB: out = out & "bin.Long2.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: cB = 1.5@: r = lA / cB: out = out & "bin.Long2.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: cB = 1.5@: r = lA \ cB: out = out & "bin.Long2.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  P011 = out
End Function
Function P012() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: lA = 2147483647: cB = 1.5@: r = lA Mod cB: out = out & "bin.Long2.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: cB = 1.5@: r = lA ^ cB: out = out & "bin.Long2.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: cB = 922337203685477.5807@: r = lA + cB: out = out & "bin.Long2.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: cB = 922337203685477.5807@: r = lA - cB: out = out & "bin.Long2.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: cB = 922337203685477.5807@: r = lA * cB: out = out & "bin.Long2.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: cB = 922337203685477.5807@: r = lA / cB: out = out & "bin.Long2.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: cB = 922337203685477.5807@: r = lA \ cB: out = out & "bin.Long2.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: cB = 922337203685477.5807@: r = lA Mod cB: out = out & "bin.Long2.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: cB = 922337203685477.5807@: r = lA ^ cB: out = out & "bin.Long2.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: vB = Empty: r = lA + vB: out = out & "bin.Long0.add.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: vB = Empty: r = lA - vB: out = out & "bin.Long0.sub.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: vB = Empty: r = lA * vB: out = out & "bin.Long0.mul.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: vB = Empty: r = lA / vB: out = out & "bin.Long0.div.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: vB = Empty: r = lA \ vB: out = out & "bin.Long0.idiv.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: vB = Empty: r = lA Mod vB: out = out & "bin.Long0.mod.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: vB = Empty: r = lA ^ vB: out = out & "bin.Long0.pow.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: vB = 7: r = lA + vB: out = out & "bin.Long0.add.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: vB = 7: r = lA - vB: out = out & "bin.Long0.sub.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: vB = 7: r = lA * vB: out = out & "bin.Long0.mul.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: vB = 7: r = lA / vB: out = out & "bin.Long0.div.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: vB = 7: r = lA \ vB: out = out & "bin.Long0.idiv.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: vB = 7: r = lA Mod vB: out = out & "bin.Long0.mod.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: vB = 7: r = lA ^ vB: out = out & "bin.Long0.pow.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: vB = 2147483647: r = lA + vB: out = out & "bin.Long0.add.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: vB = 2147483647: r = lA - vB: out = out & "bin.Long0.sub.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: vB = 2147483647: r = lA * vB: out = out & "bin.Long0.mul.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: vB = 2147483647: r = lA / vB: out = out & "bin.Long0.div.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: vB = 2147483647: r = lA \ vB: out = out & "bin.Long0.idiv.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: vB = 2147483647: r = lA Mod vB: out = out & "bin.Long0.mod.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: vB = 2147483647: r = lA ^ vB: out = out & "bin.Long0.pow.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: vB = Empty: r = lA + vB: out = out & "bin.Long1.add.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: vB = Empty: r = lA - vB: out = out & "bin.Long1.sub.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: vB = Empty: r = lA * vB: out = out & "bin.Long1.mul.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: vB = Empty: r = lA / vB: out = out & "bin.Long1.div.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: vB = Empty: r = lA \ vB: out = out & "bin.Long1.idiv.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: vB = Empty: r = lA Mod vB: out = out & "bin.Long1.mod.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: vB = Empty: r = lA ^ vB: out = out & "bin.Long1.pow.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: vB = 7: r = lA + vB: out = out & "bin.Long1.add.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: vB = 7: r = lA - vB: out = out & "bin.Long1.sub.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: vB = 7: r = lA * vB: out = out & "bin.Long1.mul.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: vB = 7: r = lA / vB: out = out & "bin.Long1.div.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: vB = 7: r = lA \ vB: out = out & "bin.Long1.idiv.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: vB = 7: r = lA Mod vB: out = out & "bin.Long1.mod.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: vB = 7: r = lA ^ vB: out = out & "bin.Long1.pow.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: vB = 2147483647: r = lA + vB: out = out & "bin.Long1.add.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: vB = 2147483647: r = lA - vB: out = out & "bin.Long1.sub.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: vB = 2147483647: r = lA * vB: out = out & "bin.Long1.mul.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: vB = 2147483647: r = lA / vB: out = out & "bin.Long1.div.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: vB = 2147483647: r = lA \ vB: out = out & "bin.Long1.idiv.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: vB = 2147483647: r = lA Mod vB: out = out & "bin.Long1.mod.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: vB = 2147483647: r = lA ^ vB: out = out & "bin.Long1.pow.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: vB = Empty: r = lA + vB: out = out & "bin.Long2.add.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: vB = Empty: r = lA - vB: out = out & "bin.Long2.sub.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: vB = Empty: r = lA * vB: out = out & "bin.Long2.mul.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: vB = Empty: r = lA / vB: out = out & "bin.Long2.div.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: vB = Empty: r = lA \ vB: out = out & "bin.Long2.idiv.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: vB = Empty: r = lA Mod vB: out = out & "bin.Long2.mod.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: vB = Empty: r = lA ^ vB: out = out & "bin.Long2.pow.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: vB = 7: r = lA + vB: out = out & "bin.Long2.add.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: vB = 7: r = lA - vB: out = out & "bin.Long2.sub.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: vB = 7: r = lA * vB: out = out & "bin.Long2.mul.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: vB = 7: r = lA / vB: out = out & "bin.Long2.div.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: vB = 7: r = lA \ vB: out = out & "bin.Long2.idiv.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: vB = 7: r = lA Mod vB: out = out & "bin.Long2.mod.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: vB = 7: r = lA ^ vB: out = out & "bin.Long2.pow.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: vB = 2147483647: r = lA + vB: out = out & "bin.Long2.add.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: vB = 2147483647: r = lA - vB: out = out & "bin.Long2.sub.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: vB = 2147483647: r = lA * vB: out = out & "bin.Long2.mul.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: vB = 2147483647: r = lA / vB: out = out & "bin.Long2.div.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: vB = 2147483647: r = lA \ vB: out = out & "bin.Long2.idiv.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: vB = 2147483647: r = lA Mod vB: out = out & "bin.Long2.mod.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: vB = 2147483647: r = lA ^ vB: out = out & "bin.Long2.pow.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: bB = 0: r = qA + bB: out = out & "bin.LongLong0.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: bB = 0: r = qA - bB: out = out & "bin.LongLong0.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: bB = 0: r = qA * bB: out = out & "bin.LongLong0.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: bB = 0: r = qA / bB: out = out & "bin.LongLong0.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: bB = 0: r = qA \ bB: out = out & "bin.LongLong0.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: bB = 0: r = qA Mod bB: out = out & "bin.LongLong0.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: bB = 0: r = qA ^ bB: out = out & "bin.LongLong0.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: bB = 200: r = qA + bB: out = out & "bin.LongLong0.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: bB = 200: r = qA - bB: out = out & "bin.LongLong0.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: bB = 200: r = qA * bB: out = out & "bin.LongLong0.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: bB = 200: r = qA / bB: out = out & "bin.LongLong0.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: bB = 200: r = qA \ bB: out = out & "bin.LongLong0.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: bB = 200: r = qA Mod bB: out = out & "bin.LongLong0.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: bB = 200: r = qA ^ bB: out = out & "bin.LongLong0.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: bB = 255: r = qA + bB: out = out & "bin.LongLong0.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: bB = 255: r = qA - bB: out = out & "bin.LongLong0.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: bB = 255: r = qA * bB: out = out & "bin.LongLong0.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: bB = 255: r = qA / bB: out = out & "bin.LongLong0.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: bB = 255: r = qA \ bB: out = out & "bin.LongLong0.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: bB = 255: r = qA Mod bB: out = out & "bin.LongLong0.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: bB = 255: r = qA ^ bB: out = out & "bin.LongLong0.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: bB = 0: r = qA + bB: out = out & "bin.LongLong1.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: bB = 0: r = qA - bB: out = out & "bin.LongLong1.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: bB = 0: r = qA * bB: out = out & "bin.LongLong1.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: bB = 0: r = qA / bB: out = out & "bin.LongLong1.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: bB = 0: r = qA \ bB: out = out & "bin.LongLong1.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: bB = 0: r = qA Mod bB: out = out & "bin.LongLong1.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: bB = 0: r = qA ^ bB: out = out & "bin.LongLong1.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: bB = 200: r = qA + bB: out = out & "bin.LongLong1.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: bB = 200: r = qA - bB: out = out & "bin.LongLong1.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: bB = 200: r = qA * bB: out = out & "bin.LongLong1.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: bB = 200: r = qA / bB: out = out & "bin.LongLong1.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: bB = 200: r = qA \ bB: out = out & "bin.LongLong1.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: bB = 200: r = qA Mod bB: out = out & "bin.LongLong1.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: bB = 200: r = qA ^ bB: out = out & "bin.LongLong1.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: bB = 255: r = qA + bB: out = out & "bin.LongLong1.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: bB = 255: r = qA - bB: out = out & "bin.LongLong1.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: bB = 255: r = qA * bB: out = out & "bin.LongLong1.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: bB = 255: r = qA / bB: out = out & "bin.LongLong1.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: bB = 255: r = qA \ bB: out = out & "bin.LongLong1.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: bB = 255: r = qA Mod bB: out = out & "bin.LongLong1.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: bB = 255: r = qA ^ bB: out = out & "bin.LongLong1.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: bB = 0: r = qA + bB: out = out & "bin.LongLong2.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: bB = 0: r = qA - bB: out = out & "bin.LongLong2.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: bB = 0: r = qA * bB: out = out & "bin.LongLong2.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: bB = 0: r = qA / bB: out = out & "bin.LongLong2.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: bB = 0: r = qA \ bB: out = out & "bin.LongLong2.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: bB = 0: r = qA Mod bB: out = out & "bin.LongLong2.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  P012 = out
End Function
Function P013() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: qA = 9223372036854775807^: bB = 0: r = qA ^ bB: out = out & "bin.LongLong2.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: bB = 200: r = qA + bB: out = out & "bin.LongLong2.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: bB = 200: r = qA - bB: out = out & "bin.LongLong2.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: bB = 200: r = qA * bB: out = out & "bin.LongLong2.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: bB = 200: r = qA / bB: out = out & "bin.LongLong2.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: bB = 200: r = qA \ bB: out = out & "bin.LongLong2.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: bB = 200: r = qA Mod bB: out = out & "bin.LongLong2.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: bB = 200: r = qA ^ bB: out = out & "bin.LongLong2.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: bB = 255: r = qA + bB: out = out & "bin.LongLong2.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: bB = 255: r = qA - bB: out = out & "bin.LongLong2.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: bB = 255: r = qA * bB: out = out & "bin.LongLong2.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: bB = 255: r = qA / bB: out = out & "bin.LongLong2.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: bB = 255: r = qA \ bB: out = out & "bin.LongLong2.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: bB = 255: r = qA Mod bB: out = out & "bin.LongLong2.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: bB = 255: r = qA ^ bB: out = out & "bin.LongLong2.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: iB = -32768: r = qA + iB: out = out & "bin.LongLong0.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: iB = -32768: r = qA - iB: out = out & "bin.LongLong0.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: iB = -32768: r = qA * iB: out = out & "bin.LongLong0.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: iB = -32768: r = qA / iB: out = out & "bin.LongLong0.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: iB = -32768: r = qA \ iB: out = out & "bin.LongLong0.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: iB = -32768: r = qA Mod iB: out = out & "bin.LongLong0.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: iB = -32768: r = qA ^ iB: out = out & "bin.LongLong0.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: iB = 1: r = qA + iB: out = out & "bin.LongLong0.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: iB = 1: r = qA - iB: out = out & "bin.LongLong0.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: iB = 1: r = qA * iB: out = out & "bin.LongLong0.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: iB = 1: r = qA / iB: out = out & "bin.LongLong0.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: iB = 1: r = qA \ iB: out = out & "bin.LongLong0.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: iB = 1: r = qA Mod iB: out = out & "bin.LongLong0.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: iB = 1: r = qA ^ iB: out = out & "bin.LongLong0.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: iB = 32767: r = qA + iB: out = out & "bin.LongLong0.add.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: iB = 32767: r = qA - iB: out = out & "bin.LongLong0.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: iB = 32767: r = qA * iB: out = out & "bin.LongLong0.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: iB = 32767: r = qA / iB: out = out & "bin.LongLong0.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: iB = 32767: r = qA \ iB: out = out & "bin.LongLong0.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: iB = 32767: r = qA Mod iB: out = out & "bin.LongLong0.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: iB = 32767: r = qA ^ iB: out = out & "bin.LongLong0.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: iB = -32768: r = qA + iB: out = out & "bin.LongLong1.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: iB = -32768: r = qA - iB: out = out & "bin.LongLong1.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: iB = -32768: r = qA * iB: out = out & "bin.LongLong1.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: iB = -32768: r = qA / iB: out = out & "bin.LongLong1.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: iB = -32768: r = qA \ iB: out = out & "bin.LongLong1.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: iB = -32768: r = qA Mod iB: out = out & "bin.LongLong1.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: iB = -32768: r = qA ^ iB: out = out & "bin.LongLong1.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: iB = 1: r = qA + iB: out = out & "bin.LongLong1.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: iB = 1: r = qA - iB: out = out & "bin.LongLong1.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: iB = 1: r = qA * iB: out = out & "bin.LongLong1.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: iB = 1: r = qA / iB: out = out & "bin.LongLong1.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: iB = 1: r = qA \ iB: out = out & "bin.LongLong1.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: iB = 1: r = qA Mod iB: out = out & "bin.LongLong1.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: iB = 1: r = qA ^ iB: out = out & "bin.LongLong1.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: iB = 32767: r = qA + iB: out = out & "bin.LongLong1.add.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: iB = 32767: r = qA - iB: out = out & "bin.LongLong1.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: iB = 32767: r = qA * iB: out = out & "bin.LongLong1.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: iB = 32767: r = qA / iB: out = out & "bin.LongLong1.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: iB = 32767: r = qA \ iB: out = out & "bin.LongLong1.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: iB = 32767: r = qA Mod iB: out = out & "bin.LongLong1.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: iB = 32767: r = qA ^ iB: out = out & "bin.LongLong1.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: iB = -32768: r = qA + iB: out = out & "bin.LongLong2.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: iB = -32768: r = qA - iB: out = out & "bin.LongLong2.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: iB = -32768: r = qA * iB: out = out & "bin.LongLong2.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: iB = -32768: r = qA / iB: out = out & "bin.LongLong2.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: iB = -32768: r = qA \ iB: out = out & "bin.LongLong2.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: iB = -32768: r = qA Mod iB: out = out & "bin.LongLong2.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: iB = -32768: r = qA ^ iB: out = out & "bin.LongLong2.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: iB = 1: r = qA + iB: out = out & "bin.LongLong2.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: iB = 1: r = qA - iB: out = out & "bin.LongLong2.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: iB = 1: r = qA * iB: out = out & "bin.LongLong2.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: iB = 1: r = qA / iB: out = out & "bin.LongLong2.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: iB = 1: r = qA \ iB: out = out & "bin.LongLong2.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: iB = 1: r = qA Mod iB: out = out & "bin.LongLong2.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: iB = 1: r = qA ^ iB: out = out & "bin.LongLong2.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: iB = 32767: r = qA + iB: out = out & "bin.LongLong2.add.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: iB = 32767: r = qA - iB: out = out & "bin.LongLong2.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: iB = 32767: r = qA * iB: out = out & "bin.LongLong2.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: iB = 32767: r = qA / iB: out = out & "bin.LongLong2.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: iB = 32767: r = qA \ iB: out = out & "bin.LongLong2.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: iB = 32767: r = qA Mod iB: out = out & "bin.LongLong2.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: iB = 32767: r = qA ^ iB: out = out & "bin.LongLong2.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: lB = -2147483648: r = qA + lB: out = out & "bin.LongLong0.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: lB = -2147483648: r = qA - lB: out = out & "bin.LongLong0.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: lB = -2147483648: r = qA * lB: out = out & "bin.LongLong0.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: lB = -2147483648: r = qA / lB: out = out & "bin.LongLong0.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: lB = -2147483648: r = qA \ lB: out = out & "bin.LongLong0.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: lB = -2147483648: r = qA Mod lB: out = out & "bin.LongLong0.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: lB = -2147483648: r = qA ^ lB: out = out & "bin.LongLong0.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: lB = 65536: r = qA + lB: out = out & "bin.LongLong0.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: lB = 65536: r = qA - lB: out = out & "bin.LongLong0.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: lB = 65536: r = qA * lB: out = out & "bin.LongLong0.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: lB = 65536: r = qA / lB: out = out & "bin.LongLong0.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: lB = 65536: r = qA \ lB: out = out & "bin.LongLong0.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: lB = 65536: r = qA Mod lB: out = out & "bin.LongLong0.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: lB = 65536: r = qA ^ lB: out = out & "bin.LongLong0.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: lB = 2147483647: r = qA + lB: out = out & "bin.LongLong0.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: lB = 2147483647: r = qA - lB: out = out & "bin.LongLong0.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: lB = 2147483647: r = qA * lB: out = out & "bin.LongLong0.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: lB = 2147483647: r = qA / lB: out = out & "bin.LongLong0.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: lB = 2147483647: r = qA \ lB: out = out & "bin.LongLong0.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: lB = 2147483647: r = qA Mod lB: out = out & "bin.LongLong0.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: lB = 2147483647: r = qA ^ lB: out = out & "bin.LongLong0.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: lB = -2147483648: r = qA + lB: out = out & "bin.LongLong1.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: lB = -2147483648: r = qA - lB: out = out & "bin.LongLong1.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: lB = -2147483648: r = qA * lB: out = out & "bin.LongLong1.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: lB = -2147483648: r = qA / lB: out = out & "bin.LongLong1.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: lB = -2147483648: r = qA \ lB: out = out & "bin.LongLong1.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: lB = -2147483648: r = qA Mod lB: out = out & "bin.LongLong1.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: lB = -2147483648: r = qA ^ lB: out = out & "bin.LongLong1.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: lB = 65536: r = qA + lB: out = out & "bin.LongLong1.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: lB = 65536: r = qA - lB: out = out & "bin.LongLong1.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: lB = 65536: r = qA * lB: out = out & "bin.LongLong1.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: lB = 65536: r = qA / lB: out = out & "bin.LongLong1.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: lB = 65536: r = qA \ lB: out = out & "bin.LongLong1.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: lB = 65536: r = qA Mod lB: out = out & "bin.LongLong1.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: lB = 65536: r = qA ^ lB: out = out & "bin.LongLong1.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: lB = 2147483647: r = qA + lB: out = out & "bin.LongLong1.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: lB = 2147483647: r = qA - lB: out = out & "bin.LongLong1.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: lB = 2147483647: r = qA * lB: out = out & "bin.LongLong1.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: lB = 2147483647: r = qA / lB: out = out & "bin.LongLong1.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: lB = 2147483647: r = qA \ lB: out = out & "bin.LongLong1.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: lB = 2147483647: r = qA Mod lB: out = out & "bin.LongLong1.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: lB = 2147483647: r = qA ^ lB: out = out & "bin.LongLong1.pow.Long2=" & Rz(Err.Number, r) & vbLf
  P013 = out
End Function
Function P014() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: qA = 9223372036854775807^: lB = -2147483648: r = qA + lB: out = out & "bin.LongLong2.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: lB = -2147483648: r = qA - lB: out = out & "bin.LongLong2.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: lB = -2147483648: r = qA * lB: out = out & "bin.LongLong2.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: lB = -2147483648: r = qA / lB: out = out & "bin.LongLong2.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: lB = -2147483648: r = qA \ lB: out = out & "bin.LongLong2.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: lB = -2147483648: r = qA Mod lB: out = out & "bin.LongLong2.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: lB = -2147483648: r = qA ^ lB: out = out & "bin.LongLong2.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: lB = 65536: r = qA + lB: out = out & "bin.LongLong2.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: lB = 65536: r = qA - lB: out = out & "bin.LongLong2.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: lB = 65536: r = qA * lB: out = out & "bin.LongLong2.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: lB = 65536: r = qA / lB: out = out & "bin.LongLong2.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: lB = 65536: r = qA \ lB: out = out & "bin.LongLong2.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: lB = 65536: r = qA Mod lB: out = out & "bin.LongLong2.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: lB = 65536: r = qA ^ lB: out = out & "bin.LongLong2.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: lB = 2147483647: r = qA + lB: out = out & "bin.LongLong2.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: lB = 2147483647: r = qA - lB: out = out & "bin.LongLong2.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: lB = 2147483647: r = qA * lB: out = out & "bin.LongLong2.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: lB = 2147483647: r = qA / lB: out = out & "bin.LongLong2.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: lB = 2147483647: r = qA \ lB: out = out & "bin.LongLong2.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: lB = 2147483647: r = qA Mod lB: out = out & "bin.LongLong2.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: lB = 2147483647: r = qA ^ lB: out = out & "bin.LongLong2.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: qB = -9223372036854775807^: r = qA + qB: out = out & "bin.LongLong0.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: qB = -9223372036854775807^: r = qA - qB: out = out & "bin.LongLong0.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: qB = -9223372036854775807^: r = qA * qB: out = out & "bin.LongLong0.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: qB = -9223372036854775807^: r = qA / qB: out = out & "bin.LongLong0.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: qB = -9223372036854775807^: r = qA \ qB: out = out & "bin.LongLong0.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: qB = -9223372036854775807^: r = qA Mod qB: out = out & "bin.LongLong0.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: qB = -9223372036854775807^: r = qA ^ qB: out = out & "bin.LongLong0.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: qB = 3: r = qA + qB: out = out & "bin.LongLong0.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: qB = 3: r = qA - qB: out = out & "bin.LongLong0.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: qB = 3: r = qA * qB: out = out & "bin.LongLong0.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: qB = 3: r = qA / qB: out = out & "bin.LongLong0.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: qB = 3: r = qA \ qB: out = out & "bin.LongLong0.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: qB = 3: r = qA Mod qB: out = out & "bin.LongLong0.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: qB = 3: r = qA ^ qB: out = out & "bin.LongLong0.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: qB = 9223372036854775807^: r = qA + qB: out = out & "bin.LongLong0.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: qB = 9223372036854775807^: r = qA - qB: out = out & "bin.LongLong0.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: qB = 9223372036854775807^: r = qA * qB: out = out & "bin.LongLong0.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: qB = 9223372036854775807^: r = qA / qB: out = out & "bin.LongLong0.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: qB = 9223372036854775807^: r = qA \ qB: out = out & "bin.LongLong0.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: qB = 9223372036854775807^: r = qA Mod qB: out = out & "bin.LongLong0.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: qB = 9223372036854775807^: r = qA ^ qB: out = out & "bin.LongLong0.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: qB = -9223372036854775807^: r = qA + qB: out = out & "bin.LongLong1.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: qB = -9223372036854775807^: r = qA - qB: out = out & "bin.LongLong1.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: qB = -9223372036854775807^: r = qA * qB: out = out & "bin.LongLong1.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: qB = -9223372036854775807^: r = qA / qB: out = out & "bin.LongLong1.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: qB = -9223372036854775807^: r = qA \ qB: out = out & "bin.LongLong1.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: qB = -9223372036854775807^: r = qA Mod qB: out = out & "bin.LongLong1.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: qB = -9223372036854775807^: r = qA ^ qB: out = out & "bin.LongLong1.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: qB = 3: r = qA + qB: out = out & "bin.LongLong1.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: qB = 3: r = qA - qB: out = out & "bin.LongLong1.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: qB = 3: r = qA * qB: out = out & "bin.LongLong1.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: qB = 3: r = qA / qB: out = out & "bin.LongLong1.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: qB = 3: r = qA \ qB: out = out & "bin.LongLong1.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: qB = 3: r = qA Mod qB: out = out & "bin.LongLong1.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: qB = 3: r = qA ^ qB: out = out & "bin.LongLong1.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: qB = 9223372036854775807^: r = qA + qB: out = out & "bin.LongLong1.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: qB = 9223372036854775807^: r = qA - qB: out = out & "bin.LongLong1.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: qB = 9223372036854775807^: r = qA * qB: out = out & "bin.LongLong1.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: qB = 9223372036854775807^: r = qA / qB: out = out & "bin.LongLong1.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: qB = 9223372036854775807^: r = qA \ qB: out = out & "bin.LongLong1.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: qB = 9223372036854775807^: r = qA Mod qB: out = out & "bin.LongLong1.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: qB = 9223372036854775807^: r = qA ^ qB: out = out & "bin.LongLong1.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: qB = -9223372036854775807^: r = qA + qB: out = out & "bin.LongLong2.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: qB = -9223372036854775807^: r = qA - qB: out = out & "bin.LongLong2.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: qB = -9223372036854775807^: r = qA * qB: out = out & "bin.LongLong2.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: qB = -9223372036854775807^: r = qA / qB: out = out & "bin.LongLong2.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: qB = -9223372036854775807^: r = qA \ qB: out = out & "bin.LongLong2.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: qB = -9223372036854775807^: r = qA Mod qB: out = out & "bin.LongLong2.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: qB = -9223372036854775807^: r = qA ^ qB: out = out & "bin.LongLong2.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: qB = 3: r = qA + qB: out = out & "bin.LongLong2.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: qB = 3: r = qA - qB: out = out & "bin.LongLong2.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: qB = 3: r = qA * qB: out = out & "bin.LongLong2.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: qB = 3: r = qA / qB: out = out & "bin.LongLong2.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: qB = 3: r = qA \ qB: out = out & "bin.LongLong2.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: qB = 3: r = qA Mod qB: out = out & "bin.LongLong2.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: qB = 3: r = qA ^ qB: out = out & "bin.LongLong2.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: qB = 9223372036854775807^: r = qA + qB: out = out & "bin.LongLong2.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: qB = 9223372036854775807^: r = qA - qB: out = out & "bin.LongLong2.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: qB = 9223372036854775807^: r = qA * qB: out = out & "bin.LongLong2.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: qB = 9223372036854775807^: r = qA / qB: out = out & "bin.LongLong2.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: qB = 9223372036854775807^: r = qA \ qB: out = out & "bin.LongLong2.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: qB = 9223372036854775807^: r = qA Mod qB: out = out & "bin.LongLong2.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: qB = 9223372036854775807^: r = qA ^ qB: out = out & "bin.LongLong2.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: sB = 0.5: r = qA + sB: out = out & "bin.LongLong0.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: sB = 0.5: r = qA - sB: out = out & "bin.LongLong0.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: sB = 0.5: r = qA * sB: out = out & "bin.LongLong0.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: sB = 0.5: r = qA / sB: out = out & "bin.LongLong0.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: sB = 0.5: r = qA \ sB: out = out & "bin.LongLong0.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: sB = 0.5: r = qA Mod sB: out = out & "bin.LongLong0.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: sB = 0.5: r = qA ^ sB: out = out & "bin.LongLong0.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: sB = 16777216: r = qA + sB: out = out & "bin.LongLong0.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: sB = 16777216: r = qA - sB: out = out & "bin.LongLong0.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: sB = 16777216: r = qA * sB: out = out & "bin.LongLong0.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: sB = 16777216: r = qA / sB: out = out & "bin.LongLong0.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: sB = 16777216: r = qA \ sB: out = out & "bin.LongLong0.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: sB = 16777216: r = qA Mod sB: out = out & "bin.LongLong0.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: sB = 16777216: r = qA ^ sB: out = out & "bin.LongLong0.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: sB = 3.0E+38: r = qA + sB: out = out & "bin.LongLong0.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: sB = 3.0E+38: r = qA - sB: out = out & "bin.LongLong0.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: sB = 3.0E+38: r = qA * sB: out = out & "bin.LongLong0.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: sB = 3.0E+38: r = qA / sB: out = out & "bin.LongLong0.div.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: sB = 3.0E+38: r = qA \ sB: out = out & "bin.LongLong0.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: sB = 3.0E+38: r = qA Mod sB: out = out & "bin.LongLong0.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: sB = 3.0E+38: r = qA ^ sB: out = out & "bin.LongLong0.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: sB = 0.5: r = qA + sB: out = out & "bin.LongLong1.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: sB = 0.5: r = qA - sB: out = out & "bin.LongLong1.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: sB = 0.5: r = qA * sB: out = out & "bin.LongLong1.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: sB = 0.5: r = qA / sB: out = out & "bin.LongLong1.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: sB = 0.5: r = qA \ sB: out = out & "bin.LongLong1.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: sB = 0.5: r = qA Mod sB: out = out & "bin.LongLong1.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: sB = 0.5: r = qA ^ sB: out = out & "bin.LongLong1.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: sB = 16777216: r = qA + sB: out = out & "bin.LongLong1.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: sB = 16777216: r = qA - sB: out = out & "bin.LongLong1.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: sB = 16777216: r = qA * sB: out = out & "bin.LongLong1.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: sB = 16777216: r = qA / sB: out = out & "bin.LongLong1.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: sB = 16777216: r = qA \ sB: out = out & "bin.LongLong1.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: sB = 16777216: r = qA Mod sB: out = out & "bin.LongLong1.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: sB = 16777216: r = qA ^ sB: out = out & "bin.LongLong1.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: sB = 3.0E+38: r = qA + sB: out = out & "bin.LongLong1.add.Single2=" & Rz(Err.Number, r) & vbLf
  P014 = out
End Function
Function P015() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: qA = 3: sB = 3.0E+38: r = qA - sB: out = out & "bin.LongLong1.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: sB = 3.0E+38: r = qA * sB: out = out & "bin.LongLong1.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: sB = 3.0E+38: r = qA / sB: out = out & "bin.LongLong1.div.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: sB = 3.0E+38: r = qA \ sB: out = out & "bin.LongLong1.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: sB = 3.0E+38: r = qA Mod sB: out = out & "bin.LongLong1.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: sB = 3.0E+38: r = qA ^ sB: out = out & "bin.LongLong1.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: sB = 0.5: r = qA + sB: out = out & "bin.LongLong2.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: sB = 0.5: r = qA - sB: out = out & "bin.LongLong2.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: sB = 0.5: r = qA * sB: out = out & "bin.LongLong2.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: sB = 0.5: r = qA / sB: out = out & "bin.LongLong2.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: sB = 0.5: r = qA \ sB: out = out & "bin.LongLong2.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: sB = 0.5: r = qA Mod sB: out = out & "bin.LongLong2.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: sB = 0.5: r = qA ^ sB: out = out & "bin.LongLong2.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: sB = 16777216: r = qA + sB: out = out & "bin.LongLong2.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: sB = 16777216: r = qA - sB: out = out & "bin.LongLong2.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: sB = 16777216: r = qA * sB: out = out & "bin.LongLong2.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: sB = 16777216: r = qA / sB: out = out & "bin.LongLong2.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: sB = 16777216: r = qA \ sB: out = out & "bin.LongLong2.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: sB = 16777216: r = qA Mod sB: out = out & "bin.LongLong2.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: sB = 16777216: r = qA ^ sB: out = out & "bin.LongLong2.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: sB = 3.0E+38: r = qA + sB: out = out & "bin.LongLong2.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: sB = 3.0E+38: r = qA - sB: out = out & "bin.LongLong2.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: sB = 3.0E+38: r = qA * sB: out = out & "bin.LongLong2.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: sB = 3.0E+38: r = qA / sB: out = out & "bin.LongLong2.div.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: sB = 3.0E+38: r = qA \ sB: out = out & "bin.LongLong2.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: sB = 3.0E+38: r = qA Mod sB: out = out & "bin.LongLong2.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: sB = 3.0E+38: r = qA ^ sB: out = out & "bin.LongLong2.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: dB = 0.5: r = qA + dB: out = out & "bin.LongLong0.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: dB = 0.5: r = qA - dB: out = out & "bin.LongLong0.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: dB = 0.5: r = qA * dB: out = out & "bin.LongLong0.mul.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: dB = 0.5: r = qA / dB: out = out & "bin.LongLong0.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: dB = 0.5: r = qA \ dB: out = out & "bin.LongLong0.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: dB = 0.5: r = qA Mod dB: out = out & "bin.LongLong0.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: dB = 0.5: r = qA ^ dB: out = out & "bin.LongLong0.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: dB = 2.5: r = qA + dB: out = out & "bin.LongLong0.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: dB = 2.5: r = qA - dB: out = out & "bin.LongLong0.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: dB = 2.5: r = qA * dB: out = out & "bin.LongLong0.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: dB = 2.5: r = qA / dB: out = out & "bin.LongLong0.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: dB = 2.5: r = qA \ dB: out = out & "bin.LongLong0.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: dB = 2.5: r = qA Mod dB: out = out & "bin.LongLong0.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: dB = 2.5: r = qA ^ dB: out = out & "bin.LongLong0.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: dB = 1.0E+308: r = qA + dB: out = out & "bin.LongLong0.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: dB = 1.0E+308: r = qA - dB: out = out & "bin.LongLong0.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: dB = 1.0E+308: r = qA * dB: out = out & "bin.LongLong0.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: dB = 1.0E+308: r = qA / dB: out = out & "bin.LongLong0.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: dB = 1.0E+308: r = qA \ dB: out = out & "bin.LongLong0.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: dB = 1.0E+308: r = qA Mod dB: out = out & "bin.LongLong0.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: dB = 1.0E+308: r = qA ^ dB: out = out & "bin.LongLong0.pow.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: dB = 0.5: r = qA + dB: out = out & "bin.LongLong1.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: dB = 0.5: r = qA - dB: out = out & "bin.LongLong1.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: dB = 0.5: r = qA * dB: out = out & "bin.LongLong1.mul.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: dB = 0.5: r = qA / dB: out = out & "bin.LongLong1.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: dB = 0.5: r = qA \ dB: out = out & "bin.LongLong1.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: dB = 0.5: r = qA Mod dB: out = out & "bin.LongLong1.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: dB = 0.5: r = qA ^ dB: out = out & "bin.LongLong1.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: dB = 2.5: r = qA + dB: out = out & "bin.LongLong1.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: dB = 2.5: r = qA - dB: out = out & "bin.LongLong1.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: dB = 2.5: r = qA * dB: out = out & "bin.LongLong1.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: dB = 2.5: r = qA / dB: out = out & "bin.LongLong1.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: dB = 2.5: r = qA \ dB: out = out & "bin.LongLong1.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: dB = 2.5: r = qA Mod dB: out = out & "bin.LongLong1.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: dB = 2.5: r = qA ^ dB: out = out & "bin.LongLong1.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: dB = 1.0E+308: r = qA + dB: out = out & "bin.LongLong1.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: dB = 1.0E+308: r = qA - dB: out = out & "bin.LongLong1.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: dB = 1.0E+308: r = qA * dB: out = out & "bin.LongLong1.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: dB = 1.0E+308: r = qA / dB: out = out & "bin.LongLong1.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: dB = 1.0E+308: r = qA \ dB: out = out & "bin.LongLong1.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: dB = 1.0E+308: r = qA Mod dB: out = out & "bin.LongLong1.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: dB = 1.0E+308: r = qA ^ dB: out = out & "bin.LongLong1.pow.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: dB = 0.5: r = qA + dB: out = out & "bin.LongLong2.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: dB = 0.5: r = qA - dB: out = out & "bin.LongLong2.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: dB = 0.5: r = qA * dB: out = out & "bin.LongLong2.mul.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: dB = 0.5: r = qA / dB: out = out & "bin.LongLong2.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: dB = 0.5: r = qA \ dB: out = out & "bin.LongLong2.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: dB = 0.5: r = qA Mod dB: out = out & "bin.LongLong2.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: dB = 0.5: r = qA ^ dB: out = out & "bin.LongLong2.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: dB = 2.5: r = qA + dB: out = out & "bin.LongLong2.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: dB = 2.5: r = qA - dB: out = out & "bin.LongLong2.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: dB = 2.5: r = qA * dB: out = out & "bin.LongLong2.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: dB = 2.5: r = qA / dB: out = out & "bin.LongLong2.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: dB = 2.5: r = qA \ dB: out = out & "bin.LongLong2.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: dB = 2.5: r = qA Mod dB: out = out & "bin.LongLong2.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: dB = 2.5: r = qA ^ dB: out = out & "bin.LongLong2.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: dB = 1.0E+308: r = qA + dB: out = out & "bin.LongLong2.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: dB = 1.0E+308: r = qA - dB: out = out & "bin.LongLong2.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: dB = 1.0E+308: r = qA * dB: out = out & "bin.LongLong2.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: dB = 1.0E+308: r = qA / dB: out = out & "bin.LongLong2.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: dB = 1.0E+308: r = qA \ dB: out = out & "bin.LongLong2.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: dB = 1.0E+308: r = qA Mod dB: out = out & "bin.LongLong2.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: dB = 1.0E+308: r = qA ^ dB: out = out & "bin.LongLong2.pow.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: cB = -0.5@: r = qA + cB: out = out & "bin.LongLong0.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: cB = -0.5@: r = qA - cB: out = out & "bin.LongLong0.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: cB = -0.5@: r = qA * cB: out = out & "bin.LongLong0.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: cB = -0.5@: r = qA / cB: out = out & "bin.LongLong0.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: cB = -0.5@: r = qA \ cB: out = out & "bin.LongLong0.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: cB = -0.5@: r = qA Mod cB: out = out & "bin.LongLong0.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: cB = -0.5@: r = qA ^ cB: out = out & "bin.LongLong0.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: cB = 1.5@: r = qA + cB: out = out & "bin.LongLong0.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: cB = 1.5@: r = qA - cB: out = out & "bin.LongLong0.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: cB = 1.5@: r = qA * cB: out = out & "bin.LongLong0.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: cB = 1.5@: r = qA / cB: out = out & "bin.LongLong0.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: cB = 1.5@: r = qA \ cB: out = out & "bin.LongLong0.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: cB = 1.5@: r = qA Mod cB: out = out & "bin.LongLong0.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: cB = 1.5@: r = qA ^ cB: out = out & "bin.LongLong0.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: cB = 922337203685477.5807@: r = qA + cB: out = out & "bin.LongLong0.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: cB = 922337203685477.5807@: r = qA - cB: out = out & "bin.LongLong0.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: cB = 922337203685477.5807@: r = qA * cB: out = out & "bin.LongLong0.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: cB = 922337203685477.5807@: r = qA / cB: out = out & "bin.LongLong0.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: cB = 922337203685477.5807@: r = qA \ cB: out = out & "bin.LongLong0.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: cB = 922337203685477.5807@: r = qA Mod cB: out = out & "bin.LongLong0.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: cB = 922337203685477.5807@: r = qA ^ cB: out = out & "bin.LongLong0.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: cB = -0.5@: r = qA + cB: out = out & "bin.LongLong1.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: cB = -0.5@: r = qA - cB: out = out & "bin.LongLong1.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: cB = -0.5@: r = qA * cB: out = out & "bin.LongLong1.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: cB = -0.5@: r = qA / cB: out = out & "bin.LongLong1.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: cB = -0.5@: r = qA \ cB: out = out & "bin.LongLong1.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: cB = -0.5@: r = qA Mod cB: out = out & "bin.LongLong1.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: cB = -0.5@: r = qA ^ cB: out = out & "bin.LongLong1.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: cB = 1.5@: r = qA + cB: out = out & "bin.LongLong1.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: cB = 1.5@: r = qA - cB: out = out & "bin.LongLong1.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  P015 = out
End Function
Function P016() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: qA = 3: cB = 1.5@: r = qA * cB: out = out & "bin.LongLong1.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: cB = 1.5@: r = qA / cB: out = out & "bin.LongLong1.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: cB = 1.5@: r = qA \ cB: out = out & "bin.LongLong1.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: cB = 1.5@: r = qA Mod cB: out = out & "bin.LongLong1.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: cB = 1.5@: r = qA ^ cB: out = out & "bin.LongLong1.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: cB = 922337203685477.5807@: r = qA + cB: out = out & "bin.LongLong1.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: cB = 922337203685477.5807@: r = qA - cB: out = out & "bin.LongLong1.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: cB = 922337203685477.5807@: r = qA * cB: out = out & "bin.LongLong1.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: cB = 922337203685477.5807@: r = qA / cB: out = out & "bin.LongLong1.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: cB = 922337203685477.5807@: r = qA \ cB: out = out & "bin.LongLong1.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: cB = 922337203685477.5807@: r = qA Mod cB: out = out & "bin.LongLong1.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: cB = 922337203685477.5807@: r = qA ^ cB: out = out & "bin.LongLong1.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: cB = -0.5@: r = qA + cB: out = out & "bin.LongLong2.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: cB = -0.5@: r = qA - cB: out = out & "bin.LongLong2.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: cB = -0.5@: r = qA * cB: out = out & "bin.LongLong2.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: cB = -0.5@: r = qA / cB: out = out & "bin.LongLong2.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: cB = -0.5@: r = qA \ cB: out = out & "bin.LongLong2.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: cB = -0.5@: r = qA Mod cB: out = out & "bin.LongLong2.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: cB = -0.5@: r = qA ^ cB: out = out & "bin.LongLong2.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: cB = 1.5@: r = qA + cB: out = out & "bin.LongLong2.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: cB = 1.5@: r = qA - cB: out = out & "bin.LongLong2.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: cB = 1.5@: r = qA * cB: out = out & "bin.LongLong2.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: cB = 1.5@: r = qA / cB: out = out & "bin.LongLong2.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: cB = 1.5@: r = qA \ cB: out = out & "bin.LongLong2.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: cB = 1.5@: r = qA Mod cB: out = out & "bin.LongLong2.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: cB = 1.5@: r = qA ^ cB: out = out & "bin.LongLong2.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: cB = 922337203685477.5807@: r = qA + cB: out = out & "bin.LongLong2.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: cB = 922337203685477.5807@: r = qA - cB: out = out & "bin.LongLong2.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: cB = 922337203685477.5807@: r = qA * cB: out = out & "bin.LongLong2.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: cB = 922337203685477.5807@: r = qA / cB: out = out & "bin.LongLong2.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: cB = 922337203685477.5807@: r = qA \ cB: out = out & "bin.LongLong2.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: cB = 922337203685477.5807@: r = qA Mod cB: out = out & "bin.LongLong2.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: cB = 922337203685477.5807@: r = qA ^ cB: out = out & "bin.LongLong2.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: vB = Empty: r = qA + vB: out = out & "bin.LongLong0.add.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: vB = Empty: r = qA - vB: out = out & "bin.LongLong0.sub.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: vB = Empty: r = qA * vB: out = out & "bin.LongLong0.mul.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: vB = Empty: r = qA / vB: out = out & "bin.LongLong0.div.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: vB = Empty: r = qA \ vB: out = out & "bin.LongLong0.idiv.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: vB = Empty: r = qA Mod vB: out = out & "bin.LongLong0.mod.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: vB = Empty: r = qA ^ vB: out = out & "bin.LongLong0.pow.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: vB = 7: r = qA + vB: out = out & "bin.LongLong0.add.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: vB = 7: r = qA - vB: out = out & "bin.LongLong0.sub.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: vB = 7: r = qA * vB: out = out & "bin.LongLong0.mul.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: vB = 7: r = qA / vB: out = out & "bin.LongLong0.div.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: vB = 7: r = qA \ vB: out = out & "bin.LongLong0.idiv.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: vB = 7: r = qA Mod vB: out = out & "bin.LongLong0.mod.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: vB = 7: r = qA ^ vB: out = out & "bin.LongLong0.pow.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: vB = 2147483647: r = qA + vB: out = out & "bin.LongLong0.add.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: vB = 2147483647: r = qA - vB: out = out & "bin.LongLong0.sub.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: vB = 2147483647: r = qA * vB: out = out & "bin.LongLong0.mul.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: vB = 2147483647: r = qA / vB: out = out & "bin.LongLong0.div.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: vB = 2147483647: r = qA \ vB: out = out & "bin.LongLong0.idiv.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: vB = 2147483647: r = qA Mod vB: out = out & "bin.LongLong0.mod.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: vB = 2147483647: r = qA ^ vB: out = out & "bin.LongLong0.pow.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: vB = Empty: r = qA + vB: out = out & "bin.LongLong1.add.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: vB = Empty: r = qA - vB: out = out & "bin.LongLong1.sub.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: vB = Empty: r = qA * vB: out = out & "bin.LongLong1.mul.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: vB = Empty: r = qA / vB: out = out & "bin.LongLong1.div.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: vB = Empty: r = qA \ vB: out = out & "bin.LongLong1.idiv.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: vB = Empty: r = qA Mod vB: out = out & "bin.LongLong1.mod.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: vB = Empty: r = qA ^ vB: out = out & "bin.LongLong1.pow.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: vB = 7: r = qA + vB: out = out & "bin.LongLong1.add.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: vB = 7: r = qA - vB: out = out & "bin.LongLong1.sub.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: vB = 7: r = qA * vB: out = out & "bin.LongLong1.mul.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: vB = 7: r = qA / vB: out = out & "bin.LongLong1.div.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: vB = 7: r = qA \ vB: out = out & "bin.LongLong1.idiv.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: vB = 7: r = qA Mod vB: out = out & "bin.LongLong1.mod.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: vB = 7: r = qA ^ vB: out = out & "bin.LongLong1.pow.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: vB = 2147483647: r = qA + vB: out = out & "bin.LongLong1.add.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: vB = 2147483647: r = qA - vB: out = out & "bin.LongLong1.sub.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: vB = 2147483647: r = qA * vB: out = out & "bin.LongLong1.mul.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: vB = 2147483647: r = qA / vB: out = out & "bin.LongLong1.div.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: vB = 2147483647: r = qA \ vB: out = out & "bin.LongLong1.idiv.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: vB = 2147483647: r = qA Mod vB: out = out & "bin.LongLong1.mod.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: vB = 2147483647: r = qA ^ vB: out = out & "bin.LongLong1.pow.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: vB = Empty: r = qA + vB: out = out & "bin.LongLong2.add.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: vB = Empty: r = qA - vB: out = out & "bin.LongLong2.sub.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: vB = Empty: r = qA * vB: out = out & "bin.LongLong2.mul.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: vB = Empty: r = qA / vB: out = out & "bin.LongLong2.div.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: vB = Empty: r = qA \ vB: out = out & "bin.LongLong2.idiv.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: vB = Empty: r = qA Mod vB: out = out & "bin.LongLong2.mod.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: vB = Empty: r = qA ^ vB: out = out & "bin.LongLong2.pow.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: vB = 7: r = qA + vB: out = out & "bin.LongLong2.add.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: vB = 7: r = qA - vB: out = out & "bin.LongLong2.sub.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: vB = 7: r = qA * vB: out = out & "bin.LongLong2.mul.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: vB = 7: r = qA / vB: out = out & "bin.LongLong2.div.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: vB = 7: r = qA \ vB: out = out & "bin.LongLong2.idiv.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: vB = 7: r = qA Mod vB: out = out & "bin.LongLong2.mod.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: vB = 7: r = qA ^ vB: out = out & "bin.LongLong2.pow.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: vB = 2147483647: r = qA + vB: out = out & "bin.LongLong2.add.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: vB = 2147483647: r = qA - vB: out = out & "bin.LongLong2.sub.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: vB = 2147483647: r = qA * vB: out = out & "bin.LongLong2.mul.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: vB = 2147483647: r = qA / vB: out = out & "bin.LongLong2.div.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: vB = 2147483647: r = qA \ vB: out = out & "bin.LongLong2.idiv.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: vB = 2147483647: r = qA Mod vB: out = out & "bin.LongLong2.mod.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: vB = 2147483647: r = qA ^ vB: out = out & "bin.LongLong2.pow.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: bB = 0: r = sA + bB: out = out & "bin.Single0.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: bB = 0: r = sA - bB: out = out & "bin.Single0.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: bB = 0: r = sA * bB: out = out & "bin.Single0.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: bB = 0: r = sA / bB: out = out & "bin.Single0.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: bB = 0: r = sA \ bB: out = out & "bin.Single0.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: bB = 0: r = sA Mod bB: out = out & "bin.Single0.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: bB = 0: r = sA ^ bB: out = out & "bin.Single0.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: bB = 200: r = sA + bB: out = out & "bin.Single0.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: bB = 200: r = sA - bB: out = out & "bin.Single0.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: bB = 200: r = sA * bB: out = out & "bin.Single0.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: bB = 200: r = sA / bB: out = out & "bin.Single0.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: bB = 200: r = sA \ bB: out = out & "bin.Single0.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: bB = 200: r = sA Mod bB: out = out & "bin.Single0.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: bB = 200: r = sA ^ bB: out = out & "bin.Single0.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: bB = 255: r = sA + bB: out = out & "bin.Single0.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: bB = 255: r = sA - bB: out = out & "bin.Single0.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: bB = 255: r = sA * bB: out = out & "bin.Single0.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: bB = 255: r = sA / bB: out = out & "bin.Single0.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: bB = 255: r = sA \ bB: out = out & "bin.Single0.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: bB = 255: r = sA Mod bB: out = out & "bin.Single0.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: bB = 255: r = sA ^ bB: out = out & "bin.Single0.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: bB = 0: r = sA + bB: out = out & "bin.Single1.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: bB = 0: r = sA - bB: out = out & "bin.Single1.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: bB = 0: r = sA * bB: out = out & "bin.Single1.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  P016 = out
End Function
Function P017() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: sA = 16777216: bB = 0: r = sA / bB: out = out & "bin.Single1.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: bB = 0: r = sA \ bB: out = out & "bin.Single1.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: bB = 0: r = sA Mod bB: out = out & "bin.Single1.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: bB = 0: r = sA ^ bB: out = out & "bin.Single1.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: bB = 200: r = sA + bB: out = out & "bin.Single1.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: bB = 200: r = sA - bB: out = out & "bin.Single1.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: bB = 200: r = sA * bB: out = out & "bin.Single1.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: bB = 200: r = sA / bB: out = out & "bin.Single1.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: bB = 200: r = sA \ bB: out = out & "bin.Single1.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: bB = 200: r = sA Mod bB: out = out & "bin.Single1.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: bB = 200: r = sA ^ bB: out = out & "bin.Single1.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: bB = 255: r = sA + bB: out = out & "bin.Single1.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: bB = 255: r = sA - bB: out = out & "bin.Single1.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: bB = 255: r = sA * bB: out = out & "bin.Single1.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: bB = 255: r = sA / bB: out = out & "bin.Single1.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: bB = 255: r = sA \ bB: out = out & "bin.Single1.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: bB = 255: r = sA Mod bB: out = out & "bin.Single1.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: bB = 255: r = sA ^ bB: out = out & "bin.Single1.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: bB = 0: r = sA + bB: out = out & "bin.Single2.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: bB = 0: r = sA - bB: out = out & "bin.Single2.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: bB = 0: r = sA * bB: out = out & "bin.Single2.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: bB = 0: r = sA / bB: out = out & "bin.Single2.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: bB = 0: r = sA \ bB: out = out & "bin.Single2.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: bB = 0: r = sA Mod bB: out = out & "bin.Single2.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: bB = 0: r = sA ^ bB: out = out & "bin.Single2.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: bB = 200: r = sA + bB: out = out & "bin.Single2.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: bB = 200: r = sA - bB: out = out & "bin.Single2.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: bB = 200: r = sA * bB: out = out & "bin.Single2.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: bB = 200: r = sA / bB: out = out & "bin.Single2.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: bB = 200: r = sA \ bB: out = out & "bin.Single2.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: bB = 200: r = sA Mod bB: out = out & "bin.Single2.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: bB = 200: r = sA ^ bB: out = out & "bin.Single2.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: bB = 255: r = sA + bB: out = out & "bin.Single2.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: bB = 255: r = sA - bB: out = out & "bin.Single2.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: bB = 255: r = sA * bB: out = out & "bin.Single2.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: bB = 255: r = sA / bB: out = out & "bin.Single2.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: bB = 255: r = sA \ bB: out = out & "bin.Single2.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: bB = 255: r = sA Mod bB: out = out & "bin.Single2.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: bB = 255: r = sA ^ bB: out = out & "bin.Single2.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: iB = -32768: r = sA + iB: out = out & "bin.Single0.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: iB = -32768: r = sA - iB: out = out & "bin.Single0.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: iB = -32768: r = sA * iB: out = out & "bin.Single0.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: iB = -32768: r = sA / iB: out = out & "bin.Single0.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: iB = -32768: r = sA \ iB: out = out & "bin.Single0.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: iB = -32768: r = sA Mod iB: out = out & "bin.Single0.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: iB = -32768: r = sA ^ iB: out = out & "bin.Single0.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: iB = 1: r = sA + iB: out = out & "bin.Single0.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: iB = 1: r = sA - iB: out = out & "bin.Single0.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: iB = 1: r = sA * iB: out = out & "bin.Single0.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: iB = 1: r = sA / iB: out = out & "bin.Single0.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: iB = 1: r = sA \ iB: out = out & "bin.Single0.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: iB = 1: r = sA Mod iB: out = out & "bin.Single0.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: iB = 1: r = sA ^ iB: out = out & "bin.Single0.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: iB = 32767: r = sA + iB: out = out & "bin.Single0.add.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: iB = 32767: r = sA - iB: out = out & "bin.Single0.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: iB = 32767: r = sA * iB: out = out & "bin.Single0.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: iB = 32767: r = sA / iB: out = out & "bin.Single0.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: iB = 32767: r = sA \ iB: out = out & "bin.Single0.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: iB = 32767: r = sA Mod iB: out = out & "bin.Single0.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: iB = 32767: r = sA ^ iB: out = out & "bin.Single0.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: iB = -32768: r = sA + iB: out = out & "bin.Single1.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: iB = -32768: r = sA - iB: out = out & "bin.Single1.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: iB = -32768: r = sA * iB: out = out & "bin.Single1.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: iB = -32768: r = sA / iB: out = out & "bin.Single1.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: iB = -32768: r = sA \ iB: out = out & "bin.Single1.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: iB = -32768: r = sA Mod iB: out = out & "bin.Single1.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: iB = -32768: r = sA ^ iB: out = out & "bin.Single1.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: iB = 1: r = sA + iB: out = out & "bin.Single1.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: iB = 1: r = sA - iB: out = out & "bin.Single1.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: iB = 1: r = sA * iB: out = out & "bin.Single1.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: iB = 1: r = sA / iB: out = out & "bin.Single1.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: iB = 1: r = sA \ iB: out = out & "bin.Single1.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: iB = 1: r = sA Mod iB: out = out & "bin.Single1.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: iB = 1: r = sA ^ iB: out = out & "bin.Single1.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: iB = 32767: r = sA + iB: out = out & "bin.Single1.add.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: iB = 32767: r = sA - iB: out = out & "bin.Single1.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: iB = 32767: r = sA * iB: out = out & "bin.Single1.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: iB = 32767: r = sA / iB: out = out & "bin.Single1.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: iB = 32767: r = sA \ iB: out = out & "bin.Single1.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: iB = 32767: r = sA Mod iB: out = out & "bin.Single1.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: iB = 32767: r = sA ^ iB: out = out & "bin.Single1.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: iB = -32768: r = sA + iB: out = out & "bin.Single2.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: iB = -32768: r = sA - iB: out = out & "bin.Single2.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: iB = -32768: r = sA * iB: out = out & "bin.Single2.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: iB = -32768: r = sA / iB: out = out & "bin.Single2.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: iB = -32768: r = sA \ iB: out = out & "bin.Single2.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: iB = -32768: r = sA Mod iB: out = out & "bin.Single2.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: iB = -32768: r = sA ^ iB: out = out & "bin.Single2.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: iB = 1: r = sA + iB: out = out & "bin.Single2.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: iB = 1: r = sA - iB: out = out & "bin.Single2.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: iB = 1: r = sA * iB: out = out & "bin.Single2.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: iB = 1: r = sA / iB: out = out & "bin.Single2.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: iB = 1: r = sA \ iB: out = out & "bin.Single2.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: iB = 1: r = sA Mod iB: out = out & "bin.Single2.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: iB = 1: r = sA ^ iB: out = out & "bin.Single2.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: iB = 32767: r = sA + iB: out = out & "bin.Single2.add.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: iB = 32767: r = sA - iB: out = out & "bin.Single2.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: iB = 32767: r = sA * iB: out = out & "bin.Single2.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: iB = 32767: r = sA / iB: out = out & "bin.Single2.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: iB = 32767: r = sA \ iB: out = out & "bin.Single2.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: iB = 32767: r = sA Mod iB: out = out & "bin.Single2.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: iB = 32767: r = sA ^ iB: out = out & "bin.Single2.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: lB = -2147483648: r = sA + lB: out = out & "bin.Single0.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: lB = -2147483648: r = sA - lB: out = out & "bin.Single0.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: lB = -2147483648: r = sA * lB: out = out & "bin.Single0.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: lB = -2147483648: r = sA / lB: out = out & "bin.Single0.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: lB = -2147483648: r = sA \ lB: out = out & "bin.Single0.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: lB = -2147483648: r = sA Mod lB: out = out & "bin.Single0.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: lB = -2147483648: r = sA ^ lB: out = out & "bin.Single0.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: lB = 65536: r = sA + lB: out = out & "bin.Single0.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: lB = 65536: r = sA - lB: out = out & "bin.Single0.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: lB = 65536: r = sA * lB: out = out & "bin.Single0.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: lB = 65536: r = sA / lB: out = out & "bin.Single0.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: lB = 65536: r = sA \ lB: out = out & "bin.Single0.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: lB = 65536: r = sA Mod lB: out = out & "bin.Single0.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: lB = 65536: r = sA ^ lB: out = out & "bin.Single0.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: lB = 2147483647: r = sA + lB: out = out & "bin.Single0.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: lB = 2147483647: r = sA - lB: out = out & "bin.Single0.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: lB = 2147483647: r = sA * lB: out = out & "bin.Single0.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: lB = 2147483647: r = sA / lB: out = out & "bin.Single0.div.Long2=" & Rz(Err.Number, r) & vbLf
  P017 = out
End Function
Function P018() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: sA = 0.5: lB = 2147483647: r = sA \ lB: out = out & "bin.Single0.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: lB = 2147483647: r = sA Mod lB: out = out & "bin.Single0.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: lB = 2147483647: r = sA ^ lB: out = out & "bin.Single0.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: lB = -2147483648: r = sA + lB: out = out & "bin.Single1.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: lB = -2147483648: r = sA - lB: out = out & "bin.Single1.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: lB = -2147483648: r = sA * lB: out = out & "bin.Single1.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: lB = -2147483648: r = sA / lB: out = out & "bin.Single1.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: lB = -2147483648: r = sA \ lB: out = out & "bin.Single1.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: lB = -2147483648: r = sA Mod lB: out = out & "bin.Single1.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: lB = -2147483648: r = sA ^ lB: out = out & "bin.Single1.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: lB = 65536: r = sA + lB: out = out & "bin.Single1.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: lB = 65536: r = sA - lB: out = out & "bin.Single1.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: lB = 65536: r = sA * lB: out = out & "bin.Single1.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: lB = 65536: r = sA / lB: out = out & "bin.Single1.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: lB = 65536: r = sA \ lB: out = out & "bin.Single1.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: lB = 65536: r = sA Mod lB: out = out & "bin.Single1.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: lB = 65536: r = sA ^ lB: out = out & "bin.Single1.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: lB = 2147483647: r = sA + lB: out = out & "bin.Single1.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: lB = 2147483647: r = sA - lB: out = out & "bin.Single1.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: lB = 2147483647: r = sA * lB: out = out & "bin.Single1.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: lB = 2147483647: r = sA / lB: out = out & "bin.Single1.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: lB = 2147483647: r = sA \ lB: out = out & "bin.Single1.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: lB = 2147483647: r = sA Mod lB: out = out & "bin.Single1.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: lB = 2147483647: r = sA ^ lB: out = out & "bin.Single1.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: lB = -2147483648: r = sA + lB: out = out & "bin.Single2.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: lB = -2147483648: r = sA - lB: out = out & "bin.Single2.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: lB = -2147483648: r = sA * lB: out = out & "bin.Single2.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: lB = -2147483648: r = sA / lB: out = out & "bin.Single2.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: lB = -2147483648: r = sA \ lB: out = out & "bin.Single2.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: lB = -2147483648: r = sA Mod lB: out = out & "bin.Single2.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: lB = -2147483648: r = sA ^ lB: out = out & "bin.Single2.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: lB = 65536: r = sA + lB: out = out & "bin.Single2.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: lB = 65536: r = sA - lB: out = out & "bin.Single2.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: lB = 65536: r = sA * lB: out = out & "bin.Single2.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: lB = 65536: r = sA / lB: out = out & "bin.Single2.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: lB = 65536: r = sA \ lB: out = out & "bin.Single2.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: lB = 65536: r = sA Mod lB: out = out & "bin.Single2.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: lB = 65536: r = sA ^ lB: out = out & "bin.Single2.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: lB = 2147483647: r = sA + lB: out = out & "bin.Single2.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: lB = 2147483647: r = sA - lB: out = out & "bin.Single2.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: lB = 2147483647: r = sA * lB: out = out & "bin.Single2.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: lB = 2147483647: r = sA / lB: out = out & "bin.Single2.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: lB = 2147483647: r = sA \ lB: out = out & "bin.Single2.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: lB = 2147483647: r = sA Mod lB: out = out & "bin.Single2.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: lB = 2147483647: r = sA ^ lB: out = out & "bin.Single2.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: qB = -9223372036854775807^: r = sA + qB: out = out & "bin.Single0.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: qB = -9223372036854775807^: r = sA - qB: out = out & "bin.Single0.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: qB = -9223372036854775807^: r = sA * qB: out = out & "bin.Single0.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: qB = -9223372036854775807^: r = sA / qB: out = out & "bin.Single0.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: qB = -9223372036854775807^: r = sA \ qB: out = out & "bin.Single0.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: qB = -9223372036854775807^: r = sA Mod qB: out = out & "bin.Single0.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: qB = -9223372036854775807^: r = sA ^ qB: out = out & "bin.Single0.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: qB = 3: r = sA + qB: out = out & "bin.Single0.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: qB = 3: r = sA - qB: out = out & "bin.Single0.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: qB = 3: r = sA * qB: out = out & "bin.Single0.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: qB = 3: r = sA / qB: out = out & "bin.Single0.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: qB = 3: r = sA \ qB: out = out & "bin.Single0.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: qB = 3: r = sA Mod qB: out = out & "bin.Single0.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: qB = 3: r = sA ^ qB: out = out & "bin.Single0.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: qB = 9223372036854775807^: r = sA + qB: out = out & "bin.Single0.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: qB = 9223372036854775807^: r = sA - qB: out = out & "bin.Single0.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: qB = 9223372036854775807^: r = sA * qB: out = out & "bin.Single0.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: qB = 9223372036854775807^: r = sA / qB: out = out & "bin.Single0.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: qB = 9223372036854775807^: r = sA \ qB: out = out & "bin.Single0.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: qB = 9223372036854775807^: r = sA Mod qB: out = out & "bin.Single0.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: qB = 9223372036854775807^: r = sA ^ qB: out = out & "bin.Single0.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: qB = -9223372036854775807^: r = sA + qB: out = out & "bin.Single1.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: qB = -9223372036854775807^: r = sA - qB: out = out & "bin.Single1.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: qB = -9223372036854775807^: r = sA * qB: out = out & "bin.Single1.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: qB = -9223372036854775807^: r = sA / qB: out = out & "bin.Single1.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: qB = -9223372036854775807^: r = sA \ qB: out = out & "bin.Single1.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: qB = -9223372036854775807^: r = sA Mod qB: out = out & "bin.Single1.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: qB = -9223372036854775807^: r = sA ^ qB: out = out & "bin.Single1.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: qB = 3: r = sA + qB: out = out & "bin.Single1.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: qB = 3: r = sA - qB: out = out & "bin.Single1.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: qB = 3: r = sA * qB: out = out & "bin.Single1.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: qB = 3: r = sA / qB: out = out & "bin.Single1.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: qB = 3: r = sA \ qB: out = out & "bin.Single1.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: qB = 3: r = sA Mod qB: out = out & "bin.Single1.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: qB = 3: r = sA ^ qB: out = out & "bin.Single1.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: qB = 9223372036854775807^: r = sA + qB: out = out & "bin.Single1.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: qB = 9223372036854775807^: r = sA - qB: out = out & "bin.Single1.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: qB = 9223372036854775807^: r = sA * qB: out = out & "bin.Single1.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: qB = 9223372036854775807^: r = sA / qB: out = out & "bin.Single1.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: qB = 9223372036854775807^: r = sA \ qB: out = out & "bin.Single1.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: qB = 9223372036854775807^: r = sA Mod qB: out = out & "bin.Single1.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: qB = 9223372036854775807^: r = sA ^ qB: out = out & "bin.Single1.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: qB = -9223372036854775807^: r = sA + qB: out = out & "bin.Single2.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: qB = -9223372036854775807^: r = sA - qB: out = out & "bin.Single2.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: qB = -9223372036854775807^: r = sA * qB: out = out & "bin.Single2.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: qB = -9223372036854775807^: r = sA / qB: out = out & "bin.Single2.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: qB = -9223372036854775807^: r = sA \ qB: out = out & "bin.Single2.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: qB = -9223372036854775807^: r = sA Mod qB: out = out & "bin.Single2.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: qB = -9223372036854775807^: r = sA ^ qB: out = out & "bin.Single2.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: qB = 3: r = sA + qB: out = out & "bin.Single2.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: qB = 3: r = sA - qB: out = out & "bin.Single2.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: qB = 3: r = sA * qB: out = out & "bin.Single2.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: qB = 3: r = sA / qB: out = out & "bin.Single2.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: qB = 3: r = sA \ qB: out = out & "bin.Single2.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: qB = 3: r = sA Mod qB: out = out & "bin.Single2.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: qB = 3: r = sA ^ qB: out = out & "bin.Single2.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: qB = 9223372036854775807^: r = sA + qB: out = out & "bin.Single2.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: qB = 9223372036854775807^: r = sA - qB: out = out & "bin.Single2.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: qB = 9223372036854775807^: r = sA * qB: out = out & "bin.Single2.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: qB = 9223372036854775807^: r = sA / qB: out = out & "bin.Single2.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: qB = 9223372036854775807^: r = sA \ qB: out = out & "bin.Single2.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: qB = 9223372036854775807^: r = sA Mod qB: out = out & "bin.Single2.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: qB = 9223372036854775807^: r = sA ^ qB: out = out & "bin.Single2.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: sB = 0.5: r = sA + sB: out = out & "bin.Single0.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: sB = 0.5: r = sA - sB: out = out & "bin.Single0.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: sB = 0.5: r = sA * sB: out = out & "bin.Single0.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: sB = 0.5: r = sA / sB: out = out & "bin.Single0.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: sB = 0.5: r = sA \ sB: out = out & "bin.Single0.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: sB = 0.5: r = sA Mod sB: out = out & "bin.Single0.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: sB = 0.5: r = sA ^ sB: out = out & "bin.Single0.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: sB = 16777216: r = sA + sB: out = out & "bin.Single0.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: sB = 16777216: r = sA - sB: out = out & "bin.Single0.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: sB = 16777216: r = sA * sB: out = out & "bin.Single0.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: sB = 16777216: r = sA / sB: out = out & "bin.Single0.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: sB = 16777216: r = sA \ sB: out = out & "bin.Single0.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  P018 = out
End Function
Function P019() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: sA = 0.5: sB = 16777216: r = sA Mod sB: out = out & "bin.Single0.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: sB = 16777216: r = sA ^ sB: out = out & "bin.Single0.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: sB = 3.0E+38: r = sA + sB: out = out & "bin.Single0.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: sB = 3.0E+38: r = sA - sB: out = out & "bin.Single0.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: sB = 3.0E+38: r = sA * sB: out = out & "bin.Single0.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: sB = 3.0E+38: r = sA / sB: out = out & "bin.Single0.div.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: sB = 3.0E+38: r = sA \ sB: out = out & "bin.Single0.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: sB = 3.0E+38: r = sA Mod sB: out = out & "bin.Single0.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: sB = 3.0E+38: r = sA ^ sB: out = out & "bin.Single0.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: sB = 0.5: r = sA + sB: out = out & "bin.Single1.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: sB = 0.5: r = sA - sB: out = out & "bin.Single1.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: sB = 0.5: r = sA * sB: out = out & "bin.Single1.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: sB = 0.5: r = sA / sB: out = out & "bin.Single1.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: sB = 0.5: r = sA \ sB: out = out & "bin.Single1.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: sB = 0.5: r = sA Mod sB: out = out & "bin.Single1.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: sB = 0.5: r = sA ^ sB: out = out & "bin.Single1.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: sB = 16777216: r = sA + sB: out = out & "bin.Single1.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: sB = 16777216: r = sA - sB: out = out & "bin.Single1.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: sB = 16777216: r = sA * sB: out = out & "bin.Single1.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: sB = 16777216: r = sA / sB: out = out & "bin.Single1.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: sB = 16777216: r = sA \ sB: out = out & "bin.Single1.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: sB = 16777216: r = sA Mod sB: out = out & "bin.Single1.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: sB = 16777216: r = sA ^ sB: out = out & "bin.Single1.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: sB = 3.0E+38: r = sA + sB: out = out & "bin.Single1.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: sB = 3.0E+38: r = sA - sB: out = out & "bin.Single1.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: sB = 3.0E+38: r = sA * sB: out = out & "bin.Single1.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: sB = 3.0E+38: r = sA / sB: out = out & "bin.Single1.div.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: sB = 3.0E+38: r = sA \ sB: out = out & "bin.Single1.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: sB = 3.0E+38: r = sA Mod sB: out = out & "bin.Single1.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: sB = 3.0E+38: r = sA ^ sB: out = out & "bin.Single1.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: sB = 0.5: r = sA + sB: out = out & "bin.Single2.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: sB = 0.5: r = sA - sB: out = out & "bin.Single2.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: sB = 0.5: r = sA * sB: out = out & "bin.Single2.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: sB = 0.5: r = sA / sB: out = out & "bin.Single2.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: sB = 0.5: r = sA \ sB: out = out & "bin.Single2.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: sB = 0.5: r = sA Mod sB: out = out & "bin.Single2.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: sB = 0.5: r = sA ^ sB: out = out & "bin.Single2.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: sB = 16777216: r = sA + sB: out = out & "bin.Single2.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: sB = 16777216: r = sA - sB: out = out & "bin.Single2.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: sB = 16777216: r = sA * sB: out = out & "bin.Single2.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: sB = 16777216: r = sA / sB: out = out & "bin.Single2.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: sB = 16777216: r = sA \ sB: out = out & "bin.Single2.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: sB = 16777216: r = sA Mod sB: out = out & "bin.Single2.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: sB = 16777216: r = sA ^ sB: out = out & "bin.Single2.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: sB = 3.0E+38: r = sA + sB: out = out & "bin.Single2.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: sB = 3.0E+38: r = sA - sB: out = out & "bin.Single2.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: sB = 3.0E+38: r = sA * sB: out = out & "bin.Single2.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: sB = 3.0E+38: r = sA / sB: out = out & "bin.Single2.div.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: sB = 3.0E+38: r = sA \ sB: out = out & "bin.Single2.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: sB = 3.0E+38: r = sA Mod sB: out = out & "bin.Single2.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: sB = 3.0E+38: r = sA ^ sB: out = out & "bin.Single2.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: dB = 0.5: r = sA + dB: out = out & "bin.Single0.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: dB = 0.5: r = sA - dB: out = out & "bin.Single0.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: dB = 0.5: r = sA * dB: out = out & "bin.Single0.mul.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: dB = 0.5: r = sA / dB: out = out & "bin.Single0.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: dB = 0.5: r = sA \ dB: out = out & "bin.Single0.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: dB = 0.5: r = sA Mod dB: out = out & "bin.Single0.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: dB = 0.5: r = sA ^ dB: out = out & "bin.Single0.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: dB = 2.5: r = sA + dB: out = out & "bin.Single0.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: dB = 2.5: r = sA - dB: out = out & "bin.Single0.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: dB = 2.5: r = sA * dB: out = out & "bin.Single0.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: dB = 2.5: r = sA / dB: out = out & "bin.Single0.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: dB = 2.5: r = sA \ dB: out = out & "bin.Single0.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: dB = 2.5: r = sA Mod dB: out = out & "bin.Single0.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: dB = 2.5: r = sA ^ dB: out = out & "bin.Single0.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: dB = 1.0E+308: r = sA + dB: out = out & "bin.Single0.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: dB = 1.0E+308: r = sA - dB: out = out & "bin.Single0.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: dB = 1.0E+308: r = sA * dB: out = out & "bin.Single0.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: dB = 1.0E+308: r = sA / dB: out = out & "bin.Single0.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: dB = 1.0E+308: r = sA \ dB: out = out & "bin.Single0.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: dB = 1.0E+308: r = sA Mod dB: out = out & "bin.Single0.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: dB = 1.0E+308: r = sA ^ dB: out = out & "bin.Single0.pow.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: dB = 0.5: r = sA + dB: out = out & "bin.Single1.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: dB = 0.5: r = sA - dB: out = out & "bin.Single1.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: dB = 0.5: r = sA * dB: out = out & "bin.Single1.mul.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: dB = 0.5: r = sA / dB: out = out & "bin.Single1.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: dB = 0.5: r = sA \ dB: out = out & "bin.Single1.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: dB = 0.5: r = sA Mod dB: out = out & "bin.Single1.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: dB = 0.5: r = sA ^ dB: out = out & "bin.Single1.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: dB = 2.5: r = sA + dB: out = out & "bin.Single1.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: dB = 2.5: r = sA - dB: out = out & "bin.Single1.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: dB = 2.5: r = sA * dB: out = out & "bin.Single1.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: dB = 2.5: r = sA / dB: out = out & "bin.Single1.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: dB = 2.5: r = sA \ dB: out = out & "bin.Single1.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: dB = 2.5: r = sA Mod dB: out = out & "bin.Single1.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: dB = 2.5: r = sA ^ dB: out = out & "bin.Single1.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: dB = 1.0E+308: r = sA + dB: out = out & "bin.Single1.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: dB = 1.0E+308: r = sA - dB: out = out & "bin.Single1.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: dB = 1.0E+308: r = sA * dB: out = out & "bin.Single1.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: dB = 1.0E+308: r = sA / dB: out = out & "bin.Single1.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: dB = 1.0E+308: r = sA \ dB: out = out & "bin.Single1.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: dB = 1.0E+308: r = sA Mod dB: out = out & "bin.Single1.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: dB = 1.0E+308: r = sA ^ dB: out = out & "bin.Single1.pow.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: dB = 0.5: r = sA + dB: out = out & "bin.Single2.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: dB = 0.5: r = sA - dB: out = out & "bin.Single2.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: dB = 0.5: r = sA * dB: out = out & "bin.Single2.mul.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: dB = 0.5: r = sA / dB: out = out & "bin.Single2.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: dB = 0.5: r = sA \ dB: out = out & "bin.Single2.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: dB = 0.5: r = sA Mod dB: out = out & "bin.Single2.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: dB = 0.5: r = sA ^ dB: out = out & "bin.Single2.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: dB = 2.5: r = sA + dB: out = out & "bin.Single2.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: dB = 2.5: r = sA - dB: out = out & "bin.Single2.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: dB = 2.5: r = sA * dB: out = out & "bin.Single2.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: dB = 2.5: r = sA / dB: out = out & "bin.Single2.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: dB = 2.5: r = sA \ dB: out = out & "bin.Single2.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: dB = 2.5: r = sA Mod dB: out = out & "bin.Single2.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: dB = 2.5: r = sA ^ dB: out = out & "bin.Single2.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: dB = 1.0E+308: r = sA + dB: out = out & "bin.Single2.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: dB = 1.0E+308: r = sA - dB: out = out & "bin.Single2.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: dB = 1.0E+308: r = sA * dB: out = out & "bin.Single2.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: dB = 1.0E+308: r = sA / dB: out = out & "bin.Single2.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: dB = 1.0E+308: r = sA \ dB: out = out & "bin.Single2.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: dB = 1.0E+308: r = sA Mod dB: out = out & "bin.Single2.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: dB = 1.0E+308: r = sA ^ dB: out = out & "bin.Single2.pow.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: cB = -0.5@: r = sA + cB: out = out & "bin.Single0.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: cB = -0.5@: r = sA - cB: out = out & "bin.Single0.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: cB = -0.5@: r = sA * cB: out = out & "bin.Single0.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: cB = -0.5@: r = sA / cB: out = out & "bin.Single0.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: cB = -0.5@: r = sA \ cB: out = out & "bin.Single0.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: cB = -0.5@: r = sA Mod cB: out = out & "bin.Single0.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  P019 = out
End Function
Function P020() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: sA = 0.5: cB = -0.5@: r = sA ^ cB: out = out & "bin.Single0.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: cB = 1.5@: r = sA + cB: out = out & "bin.Single0.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: cB = 1.5@: r = sA - cB: out = out & "bin.Single0.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: cB = 1.5@: r = sA * cB: out = out & "bin.Single0.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: cB = 1.5@: r = sA / cB: out = out & "bin.Single0.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: cB = 1.5@: r = sA \ cB: out = out & "bin.Single0.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: cB = 1.5@: r = sA Mod cB: out = out & "bin.Single0.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: cB = 1.5@: r = sA ^ cB: out = out & "bin.Single0.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: cB = 922337203685477.5807@: r = sA + cB: out = out & "bin.Single0.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: cB = 922337203685477.5807@: r = sA - cB: out = out & "bin.Single0.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: cB = 922337203685477.5807@: r = sA * cB: out = out & "bin.Single0.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: cB = 922337203685477.5807@: r = sA / cB: out = out & "bin.Single0.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: cB = 922337203685477.5807@: r = sA \ cB: out = out & "bin.Single0.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: cB = 922337203685477.5807@: r = sA Mod cB: out = out & "bin.Single0.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: cB = 922337203685477.5807@: r = sA ^ cB: out = out & "bin.Single0.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: cB = -0.5@: r = sA + cB: out = out & "bin.Single1.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: cB = -0.5@: r = sA - cB: out = out & "bin.Single1.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: cB = -0.5@: r = sA * cB: out = out & "bin.Single1.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: cB = -0.5@: r = sA / cB: out = out & "bin.Single1.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: cB = -0.5@: r = sA \ cB: out = out & "bin.Single1.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: cB = -0.5@: r = sA Mod cB: out = out & "bin.Single1.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: cB = -0.5@: r = sA ^ cB: out = out & "bin.Single1.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: cB = 1.5@: r = sA + cB: out = out & "bin.Single1.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: cB = 1.5@: r = sA - cB: out = out & "bin.Single1.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: cB = 1.5@: r = sA * cB: out = out & "bin.Single1.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: cB = 1.5@: r = sA / cB: out = out & "bin.Single1.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: cB = 1.5@: r = sA \ cB: out = out & "bin.Single1.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: cB = 1.5@: r = sA Mod cB: out = out & "bin.Single1.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: cB = 1.5@: r = sA ^ cB: out = out & "bin.Single1.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: cB = 922337203685477.5807@: r = sA + cB: out = out & "bin.Single1.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: cB = 922337203685477.5807@: r = sA - cB: out = out & "bin.Single1.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: cB = 922337203685477.5807@: r = sA * cB: out = out & "bin.Single1.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: cB = 922337203685477.5807@: r = sA / cB: out = out & "bin.Single1.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: cB = 922337203685477.5807@: r = sA \ cB: out = out & "bin.Single1.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: cB = 922337203685477.5807@: r = sA Mod cB: out = out & "bin.Single1.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: cB = 922337203685477.5807@: r = sA ^ cB: out = out & "bin.Single1.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: cB = -0.5@: r = sA + cB: out = out & "bin.Single2.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: cB = -0.5@: r = sA - cB: out = out & "bin.Single2.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: cB = -0.5@: r = sA * cB: out = out & "bin.Single2.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: cB = -0.5@: r = sA / cB: out = out & "bin.Single2.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: cB = -0.5@: r = sA \ cB: out = out & "bin.Single2.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: cB = -0.5@: r = sA Mod cB: out = out & "bin.Single2.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: cB = -0.5@: r = sA ^ cB: out = out & "bin.Single2.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: cB = 1.5@: r = sA + cB: out = out & "bin.Single2.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: cB = 1.5@: r = sA - cB: out = out & "bin.Single2.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: cB = 1.5@: r = sA * cB: out = out & "bin.Single2.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: cB = 1.5@: r = sA / cB: out = out & "bin.Single2.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: cB = 1.5@: r = sA \ cB: out = out & "bin.Single2.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: cB = 1.5@: r = sA Mod cB: out = out & "bin.Single2.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: cB = 1.5@: r = sA ^ cB: out = out & "bin.Single2.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: cB = 922337203685477.5807@: r = sA + cB: out = out & "bin.Single2.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: cB = 922337203685477.5807@: r = sA - cB: out = out & "bin.Single2.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: cB = 922337203685477.5807@: r = sA * cB: out = out & "bin.Single2.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: cB = 922337203685477.5807@: r = sA / cB: out = out & "bin.Single2.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: cB = 922337203685477.5807@: r = sA \ cB: out = out & "bin.Single2.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: cB = 922337203685477.5807@: r = sA Mod cB: out = out & "bin.Single2.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: cB = 922337203685477.5807@: r = sA ^ cB: out = out & "bin.Single2.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: vB = Empty: r = sA + vB: out = out & "bin.Single0.add.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: vB = Empty: r = sA - vB: out = out & "bin.Single0.sub.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: vB = Empty: r = sA * vB: out = out & "bin.Single0.mul.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: vB = Empty: r = sA / vB: out = out & "bin.Single0.div.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: vB = Empty: r = sA \ vB: out = out & "bin.Single0.idiv.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: vB = Empty: r = sA Mod vB: out = out & "bin.Single0.mod.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: vB = Empty: r = sA ^ vB: out = out & "bin.Single0.pow.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: vB = 7: r = sA + vB: out = out & "bin.Single0.add.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: vB = 7: r = sA - vB: out = out & "bin.Single0.sub.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: vB = 7: r = sA * vB: out = out & "bin.Single0.mul.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: vB = 7: r = sA / vB: out = out & "bin.Single0.div.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: vB = 7: r = sA \ vB: out = out & "bin.Single0.idiv.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: vB = 7: r = sA Mod vB: out = out & "bin.Single0.mod.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: vB = 7: r = sA ^ vB: out = out & "bin.Single0.pow.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: vB = 2147483647: r = sA + vB: out = out & "bin.Single0.add.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: vB = 2147483647: r = sA - vB: out = out & "bin.Single0.sub.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: vB = 2147483647: r = sA * vB: out = out & "bin.Single0.mul.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: vB = 2147483647: r = sA / vB: out = out & "bin.Single0.div.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: vB = 2147483647: r = sA \ vB: out = out & "bin.Single0.idiv.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: vB = 2147483647: r = sA Mod vB: out = out & "bin.Single0.mod.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: vB = 2147483647: r = sA ^ vB: out = out & "bin.Single0.pow.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: vB = Empty: r = sA + vB: out = out & "bin.Single1.add.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: vB = Empty: r = sA - vB: out = out & "bin.Single1.sub.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: vB = Empty: r = sA * vB: out = out & "bin.Single1.mul.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: vB = Empty: r = sA / vB: out = out & "bin.Single1.div.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: vB = Empty: r = sA \ vB: out = out & "bin.Single1.idiv.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: vB = Empty: r = sA Mod vB: out = out & "bin.Single1.mod.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: vB = Empty: r = sA ^ vB: out = out & "bin.Single1.pow.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: vB = 7: r = sA + vB: out = out & "bin.Single1.add.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: vB = 7: r = sA - vB: out = out & "bin.Single1.sub.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: vB = 7: r = sA * vB: out = out & "bin.Single1.mul.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: vB = 7: r = sA / vB: out = out & "bin.Single1.div.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: vB = 7: r = sA \ vB: out = out & "bin.Single1.idiv.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: vB = 7: r = sA Mod vB: out = out & "bin.Single1.mod.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: vB = 7: r = sA ^ vB: out = out & "bin.Single1.pow.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: vB = 2147483647: r = sA + vB: out = out & "bin.Single1.add.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: vB = 2147483647: r = sA - vB: out = out & "bin.Single1.sub.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: vB = 2147483647: r = sA * vB: out = out & "bin.Single1.mul.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: vB = 2147483647: r = sA / vB: out = out & "bin.Single1.div.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: vB = 2147483647: r = sA \ vB: out = out & "bin.Single1.idiv.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: vB = 2147483647: r = sA Mod vB: out = out & "bin.Single1.mod.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: vB = 2147483647: r = sA ^ vB: out = out & "bin.Single1.pow.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: vB = Empty: r = sA + vB: out = out & "bin.Single2.add.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: vB = Empty: r = sA - vB: out = out & "bin.Single2.sub.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: vB = Empty: r = sA * vB: out = out & "bin.Single2.mul.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: vB = Empty: r = sA / vB: out = out & "bin.Single2.div.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: vB = Empty: r = sA \ vB: out = out & "bin.Single2.idiv.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: vB = Empty: r = sA Mod vB: out = out & "bin.Single2.mod.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: vB = Empty: r = sA ^ vB: out = out & "bin.Single2.pow.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: vB = 7: r = sA + vB: out = out & "bin.Single2.add.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: vB = 7: r = sA - vB: out = out & "bin.Single2.sub.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: vB = 7: r = sA * vB: out = out & "bin.Single2.mul.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: vB = 7: r = sA / vB: out = out & "bin.Single2.div.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: vB = 7: r = sA \ vB: out = out & "bin.Single2.idiv.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: vB = 7: r = sA Mod vB: out = out & "bin.Single2.mod.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: vB = 7: r = sA ^ vB: out = out & "bin.Single2.pow.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: vB = 2147483647: r = sA + vB: out = out & "bin.Single2.add.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: vB = 2147483647: r = sA - vB: out = out & "bin.Single2.sub.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: vB = 2147483647: r = sA * vB: out = out & "bin.Single2.mul.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: vB = 2147483647: r = sA / vB: out = out & "bin.Single2.div.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: vB = 2147483647: r = sA \ vB: out = out & "bin.Single2.idiv.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: vB = 2147483647: r = sA Mod vB: out = out & "bin.Single2.mod.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: vB = 2147483647: r = sA ^ vB: out = out & "bin.Single2.pow.Variant2=" & Rz(Err.Number, r) & vbLf
  P020 = out
End Function
Function P021() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: dA = 0.5: bB = 0: r = dA + bB: out = out & "bin.Double0.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: bB = 0: r = dA - bB: out = out & "bin.Double0.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: bB = 0: r = dA * bB: out = out & "bin.Double0.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: bB = 0: r = dA / bB: out = out & "bin.Double0.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: bB = 0: r = dA \ bB: out = out & "bin.Double0.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: bB = 0: r = dA Mod bB: out = out & "bin.Double0.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: bB = 0: r = dA ^ bB: out = out & "bin.Double0.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: bB = 200: r = dA + bB: out = out & "bin.Double0.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: bB = 200: r = dA - bB: out = out & "bin.Double0.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: bB = 200: r = dA * bB: out = out & "bin.Double0.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: bB = 200: r = dA / bB: out = out & "bin.Double0.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: bB = 200: r = dA \ bB: out = out & "bin.Double0.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: bB = 200: r = dA Mod bB: out = out & "bin.Double0.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: bB = 200: r = dA ^ bB: out = out & "bin.Double0.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: bB = 255: r = dA + bB: out = out & "bin.Double0.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: bB = 255: r = dA - bB: out = out & "bin.Double0.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: bB = 255: r = dA * bB: out = out & "bin.Double0.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: bB = 255: r = dA / bB: out = out & "bin.Double0.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: bB = 255: r = dA \ bB: out = out & "bin.Double0.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: bB = 255: r = dA Mod bB: out = out & "bin.Double0.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: bB = 255: r = dA ^ bB: out = out & "bin.Double0.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: bB = 0: r = dA + bB: out = out & "bin.Double1.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: bB = 0: r = dA - bB: out = out & "bin.Double1.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: bB = 0: r = dA * bB: out = out & "bin.Double1.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: bB = 0: r = dA / bB: out = out & "bin.Double1.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: bB = 0: r = dA \ bB: out = out & "bin.Double1.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: bB = 0: r = dA Mod bB: out = out & "bin.Double1.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: bB = 0: r = dA ^ bB: out = out & "bin.Double1.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: bB = 200: r = dA + bB: out = out & "bin.Double1.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: bB = 200: r = dA - bB: out = out & "bin.Double1.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: bB = 200: r = dA * bB: out = out & "bin.Double1.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: bB = 200: r = dA / bB: out = out & "bin.Double1.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: bB = 200: r = dA \ bB: out = out & "bin.Double1.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: bB = 200: r = dA Mod bB: out = out & "bin.Double1.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: bB = 200: r = dA ^ bB: out = out & "bin.Double1.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: bB = 255: r = dA + bB: out = out & "bin.Double1.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: bB = 255: r = dA - bB: out = out & "bin.Double1.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: bB = 255: r = dA * bB: out = out & "bin.Double1.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: bB = 255: r = dA / bB: out = out & "bin.Double1.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: bB = 255: r = dA \ bB: out = out & "bin.Double1.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: bB = 255: r = dA Mod bB: out = out & "bin.Double1.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: bB = 255: r = dA ^ bB: out = out & "bin.Double1.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: bB = 0: r = dA + bB: out = out & "bin.Double2.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: bB = 0: r = dA - bB: out = out & "bin.Double2.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: bB = 0: r = dA * bB: out = out & "bin.Double2.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: bB = 0: r = dA / bB: out = out & "bin.Double2.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: bB = 0: r = dA \ bB: out = out & "bin.Double2.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: bB = 0: r = dA Mod bB: out = out & "bin.Double2.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: bB = 0: r = dA ^ bB: out = out & "bin.Double2.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: bB = 200: r = dA + bB: out = out & "bin.Double2.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: bB = 200: r = dA - bB: out = out & "bin.Double2.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: bB = 200: r = dA * bB: out = out & "bin.Double2.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: bB = 200: r = dA / bB: out = out & "bin.Double2.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: bB = 200: r = dA \ bB: out = out & "bin.Double2.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: bB = 200: r = dA Mod bB: out = out & "bin.Double2.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: bB = 200: r = dA ^ bB: out = out & "bin.Double2.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: bB = 255: r = dA + bB: out = out & "bin.Double2.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: bB = 255: r = dA - bB: out = out & "bin.Double2.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: bB = 255: r = dA * bB: out = out & "bin.Double2.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: bB = 255: r = dA / bB: out = out & "bin.Double2.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: bB = 255: r = dA \ bB: out = out & "bin.Double2.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: bB = 255: r = dA Mod bB: out = out & "bin.Double2.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: bB = 255: r = dA ^ bB: out = out & "bin.Double2.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: iB = -32768: r = dA + iB: out = out & "bin.Double0.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: iB = -32768: r = dA - iB: out = out & "bin.Double0.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: iB = -32768: r = dA * iB: out = out & "bin.Double0.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: iB = -32768: r = dA / iB: out = out & "bin.Double0.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: iB = -32768: r = dA \ iB: out = out & "bin.Double0.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: iB = -32768: r = dA Mod iB: out = out & "bin.Double0.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: iB = -32768: r = dA ^ iB: out = out & "bin.Double0.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: iB = 1: r = dA + iB: out = out & "bin.Double0.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: iB = 1: r = dA - iB: out = out & "bin.Double0.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: iB = 1: r = dA * iB: out = out & "bin.Double0.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: iB = 1: r = dA / iB: out = out & "bin.Double0.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: iB = 1: r = dA \ iB: out = out & "bin.Double0.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: iB = 1: r = dA Mod iB: out = out & "bin.Double0.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: iB = 1: r = dA ^ iB: out = out & "bin.Double0.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: iB = 32767: r = dA + iB: out = out & "bin.Double0.add.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: iB = 32767: r = dA - iB: out = out & "bin.Double0.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: iB = 32767: r = dA * iB: out = out & "bin.Double0.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: iB = 32767: r = dA / iB: out = out & "bin.Double0.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: iB = 32767: r = dA \ iB: out = out & "bin.Double0.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: iB = 32767: r = dA Mod iB: out = out & "bin.Double0.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: iB = 32767: r = dA ^ iB: out = out & "bin.Double0.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: iB = -32768: r = dA + iB: out = out & "bin.Double1.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: iB = -32768: r = dA - iB: out = out & "bin.Double1.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: iB = -32768: r = dA * iB: out = out & "bin.Double1.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: iB = -32768: r = dA / iB: out = out & "bin.Double1.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: iB = -32768: r = dA \ iB: out = out & "bin.Double1.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: iB = -32768: r = dA Mod iB: out = out & "bin.Double1.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: iB = -32768: r = dA ^ iB: out = out & "bin.Double1.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: iB = 1: r = dA + iB: out = out & "bin.Double1.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: iB = 1: r = dA - iB: out = out & "bin.Double1.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: iB = 1: r = dA * iB: out = out & "bin.Double1.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: iB = 1: r = dA / iB: out = out & "bin.Double1.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: iB = 1: r = dA \ iB: out = out & "bin.Double1.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: iB = 1: r = dA Mod iB: out = out & "bin.Double1.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: iB = 1: r = dA ^ iB: out = out & "bin.Double1.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: iB = 32767: r = dA + iB: out = out & "bin.Double1.add.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: iB = 32767: r = dA - iB: out = out & "bin.Double1.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: iB = 32767: r = dA * iB: out = out & "bin.Double1.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: iB = 32767: r = dA / iB: out = out & "bin.Double1.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: iB = 32767: r = dA \ iB: out = out & "bin.Double1.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: iB = 32767: r = dA Mod iB: out = out & "bin.Double1.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: iB = 32767: r = dA ^ iB: out = out & "bin.Double1.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: iB = -32768: r = dA + iB: out = out & "bin.Double2.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: iB = -32768: r = dA - iB: out = out & "bin.Double2.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: iB = -32768: r = dA * iB: out = out & "bin.Double2.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: iB = -32768: r = dA / iB: out = out & "bin.Double2.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: iB = -32768: r = dA \ iB: out = out & "bin.Double2.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: iB = -32768: r = dA Mod iB: out = out & "bin.Double2.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: iB = -32768: r = dA ^ iB: out = out & "bin.Double2.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: iB = 1: r = dA + iB: out = out & "bin.Double2.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: iB = 1: r = dA - iB: out = out & "bin.Double2.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: iB = 1: r = dA * iB: out = out & "bin.Double2.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: iB = 1: r = dA / iB: out = out & "bin.Double2.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: iB = 1: r = dA \ iB: out = out & "bin.Double2.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: iB = 1: r = dA Mod iB: out = out & "bin.Double2.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: iB = 1: r = dA ^ iB: out = out & "bin.Double2.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: iB = 32767: r = dA + iB: out = out & "bin.Double2.add.Integer2=" & Rz(Err.Number, r) & vbLf
  P021 = out
End Function
Function P022() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: dA = 1.0E+308: iB = 32767: r = dA - iB: out = out & "bin.Double2.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: iB = 32767: r = dA * iB: out = out & "bin.Double2.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: iB = 32767: r = dA / iB: out = out & "bin.Double2.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: iB = 32767: r = dA \ iB: out = out & "bin.Double2.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: iB = 32767: r = dA Mod iB: out = out & "bin.Double2.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: iB = 32767: r = dA ^ iB: out = out & "bin.Double2.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: lB = -2147483648: r = dA + lB: out = out & "bin.Double0.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: lB = -2147483648: r = dA - lB: out = out & "bin.Double0.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: lB = -2147483648: r = dA * lB: out = out & "bin.Double0.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: lB = -2147483648: r = dA / lB: out = out & "bin.Double0.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: lB = -2147483648: r = dA \ lB: out = out & "bin.Double0.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: lB = -2147483648: r = dA Mod lB: out = out & "bin.Double0.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: lB = -2147483648: r = dA ^ lB: out = out & "bin.Double0.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: lB = 65536: r = dA + lB: out = out & "bin.Double0.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: lB = 65536: r = dA - lB: out = out & "bin.Double0.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: lB = 65536: r = dA * lB: out = out & "bin.Double0.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: lB = 65536: r = dA / lB: out = out & "bin.Double0.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: lB = 65536: r = dA \ lB: out = out & "bin.Double0.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: lB = 65536: r = dA Mod lB: out = out & "bin.Double0.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: lB = 65536: r = dA ^ lB: out = out & "bin.Double0.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: lB = 2147483647: r = dA + lB: out = out & "bin.Double0.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: lB = 2147483647: r = dA - lB: out = out & "bin.Double0.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: lB = 2147483647: r = dA * lB: out = out & "bin.Double0.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: lB = 2147483647: r = dA / lB: out = out & "bin.Double0.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: lB = 2147483647: r = dA \ lB: out = out & "bin.Double0.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: lB = 2147483647: r = dA Mod lB: out = out & "bin.Double0.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: lB = 2147483647: r = dA ^ lB: out = out & "bin.Double0.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: lB = -2147483648: r = dA + lB: out = out & "bin.Double1.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: lB = -2147483648: r = dA - lB: out = out & "bin.Double1.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: lB = -2147483648: r = dA * lB: out = out & "bin.Double1.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: lB = -2147483648: r = dA / lB: out = out & "bin.Double1.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: lB = -2147483648: r = dA \ lB: out = out & "bin.Double1.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: lB = -2147483648: r = dA Mod lB: out = out & "bin.Double1.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: lB = -2147483648: r = dA ^ lB: out = out & "bin.Double1.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: lB = 65536: r = dA + lB: out = out & "bin.Double1.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: lB = 65536: r = dA - lB: out = out & "bin.Double1.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: lB = 65536: r = dA * lB: out = out & "bin.Double1.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: lB = 65536: r = dA / lB: out = out & "bin.Double1.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: lB = 65536: r = dA \ lB: out = out & "bin.Double1.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: lB = 65536: r = dA Mod lB: out = out & "bin.Double1.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: lB = 65536: r = dA ^ lB: out = out & "bin.Double1.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: lB = 2147483647: r = dA + lB: out = out & "bin.Double1.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: lB = 2147483647: r = dA - lB: out = out & "bin.Double1.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: lB = 2147483647: r = dA * lB: out = out & "bin.Double1.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: lB = 2147483647: r = dA / lB: out = out & "bin.Double1.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: lB = 2147483647: r = dA \ lB: out = out & "bin.Double1.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: lB = 2147483647: r = dA Mod lB: out = out & "bin.Double1.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: lB = 2147483647: r = dA ^ lB: out = out & "bin.Double1.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: lB = -2147483648: r = dA + lB: out = out & "bin.Double2.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: lB = -2147483648: r = dA - lB: out = out & "bin.Double2.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: lB = -2147483648: r = dA * lB: out = out & "bin.Double2.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: lB = -2147483648: r = dA / lB: out = out & "bin.Double2.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: lB = -2147483648: r = dA \ lB: out = out & "bin.Double2.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: lB = -2147483648: r = dA Mod lB: out = out & "bin.Double2.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: lB = -2147483648: r = dA ^ lB: out = out & "bin.Double2.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: lB = 65536: r = dA + lB: out = out & "bin.Double2.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: lB = 65536: r = dA - lB: out = out & "bin.Double2.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: lB = 65536: r = dA * lB: out = out & "bin.Double2.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: lB = 65536: r = dA / lB: out = out & "bin.Double2.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: lB = 65536: r = dA \ lB: out = out & "bin.Double2.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: lB = 65536: r = dA Mod lB: out = out & "bin.Double2.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: lB = 65536: r = dA ^ lB: out = out & "bin.Double2.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: lB = 2147483647: r = dA + lB: out = out & "bin.Double2.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: lB = 2147483647: r = dA - lB: out = out & "bin.Double2.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: lB = 2147483647: r = dA * lB: out = out & "bin.Double2.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: lB = 2147483647: r = dA / lB: out = out & "bin.Double2.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: lB = 2147483647: r = dA \ lB: out = out & "bin.Double2.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: lB = 2147483647: r = dA Mod lB: out = out & "bin.Double2.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: lB = 2147483647: r = dA ^ lB: out = out & "bin.Double2.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: qB = -9223372036854775807^: r = dA + qB: out = out & "bin.Double0.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: qB = -9223372036854775807^: r = dA - qB: out = out & "bin.Double0.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: qB = -9223372036854775807^: r = dA * qB: out = out & "bin.Double0.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: qB = -9223372036854775807^: r = dA / qB: out = out & "bin.Double0.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: qB = -9223372036854775807^: r = dA \ qB: out = out & "bin.Double0.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: qB = -9223372036854775807^: r = dA Mod qB: out = out & "bin.Double0.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: qB = -9223372036854775807^: r = dA ^ qB: out = out & "bin.Double0.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: qB = 3: r = dA + qB: out = out & "bin.Double0.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: qB = 3: r = dA - qB: out = out & "bin.Double0.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: qB = 3: r = dA * qB: out = out & "bin.Double0.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: qB = 3: r = dA / qB: out = out & "bin.Double0.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: qB = 3: r = dA \ qB: out = out & "bin.Double0.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: qB = 3: r = dA Mod qB: out = out & "bin.Double0.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: qB = 3: r = dA ^ qB: out = out & "bin.Double0.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: qB = 9223372036854775807^: r = dA + qB: out = out & "bin.Double0.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: qB = 9223372036854775807^: r = dA - qB: out = out & "bin.Double0.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: qB = 9223372036854775807^: r = dA * qB: out = out & "bin.Double0.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: qB = 9223372036854775807^: r = dA / qB: out = out & "bin.Double0.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: qB = 9223372036854775807^: r = dA \ qB: out = out & "bin.Double0.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: qB = 9223372036854775807^: r = dA Mod qB: out = out & "bin.Double0.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: qB = 9223372036854775807^: r = dA ^ qB: out = out & "bin.Double0.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: qB = -9223372036854775807^: r = dA + qB: out = out & "bin.Double1.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: qB = -9223372036854775807^: r = dA - qB: out = out & "bin.Double1.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: qB = -9223372036854775807^: r = dA * qB: out = out & "bin.Double1.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: qB = -9223372036854775807^: r = dA / qB: out = out & "bin.Double1.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: qB = -9223372036854775807^: r = dA \ qB: out = out & "bin.Double1.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: qB = -9223372036854775807^: r = dA Mod qB: out = out & "bin.Double1.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: qB = -9223372036854775807^: r = dA ^ qB: out = out & "bin.Double1.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: qB = 3: r = dA + qB: out = out & "bin.Double1.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: qB = 3: r = dA - qB: out = out & "bin.Double1.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: qB = 3: r = dA * qB: out = out & "bin.Double1.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: qB = 3: r = dA / qB: out = out & "bin.Double1.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: qB = 3: r = dA \ qB: out = out & "bin.Double1.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: qB = 3: r = dA Mod qB: out = out & "bin.Double1.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: qB = 3: r = dA ^ qB: out = out & "bin.Double1.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: qB = 9223372036854775807^: r = dA + qB: out = out & "bin.Double1.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: qB = 9223372036854775807^: r = dA - qB: out = out & "bin.Double1.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: qB = 9223372036854775807^: r = dA * qB: out = out & "bin.Double1.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: qB = 9223372036854775807^: r = dA / qB: out = out & "bin.Double1.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: qB = 9223372036854775807^: r = dA \ qB: out = out & "bin.Double1.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: qB = 9223372036854775807^: r = dA Mod qB: out = out & "bin.Double1.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: qB = 9223372036854775807^: r = dA ^ qB: out = out & "bin.Double1.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: qB = -9223372036854775807^: r = dA + qB: out = out & "bin.Double2.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: qB = -9223372036854775807^: r = dA - qB: out = out & "bin.Double2.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: qB = -9223372036854775807^: r = dA * qB: out = out & "bin.Double2.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: qB = -9223372036854775807^: r = dA / qB: out = out & "bin.Double2.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: qB = -9223372036854775807^: r = dA \ qB: out = out & "bin.Double2.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: qB = -9223372036854775807^: r = dA Mod qB: out = out & "bin.Double2.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: qB = -9223372036854775807^: r = dA ^ qB: out = out & "bin.Double2.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: qB = 3: r = dA + qB: out = out & "bin.Double2.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: qB = 3: r = dA - qB: out = out & "bin.Double2.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  P022 = out
End Function
Function P023() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: dA = 1.0E+308: qB = 3: r = dA * qB: out = out & "bin.Double2.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: qB = 3: r = dA / qB: out = out & "bin.Double2.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: qB = 3: r = dA \ qB: out = out & "bin.Double2.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: qB = 3: r = dA Mod qB: out = out & "bin.Double2.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: qB = 3: r = dA ^ qB: out = out & "bin.Double2.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: qB = 9223372036854775807^: r = dA + qB: out = out & "bin.Double2.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: qB = 9223372036854775807^: r = dA - qB: out = out & "bin.Double2.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: qB = 9223372036854775807^: r = dA * qB: out = out & "bin.Double2.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: qB = 9223372036854775807^: r = dA / qB: out = out & "bin.Double2.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: qB = 9223372036854775807^: r = dA \ qB: out = out & "bin.Double2.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: qB = 9223372036854775807^: r = dA Mod qB: out = out & "bin.Double2.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: qB = 9223372036854775807^: r = dA ^ qB: out = out & "bin.Double2.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: sB = 0.5: r = dA + sB: out = out & "bin.Double0.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: sB = 0.5: r = dA - sB: out = out & "bin.Double0.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: sB = 0.5: r = dA * sB: out = out & "bin.Double0.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: sB = 0.5: r = dA / sB: out = out & "bin.Double0.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: sB = 0.5: r = dA \ sB: out = out & "bin.Double0.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: sB = 0.5: r = dA Mod sB: out = out & "bin.Double0.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: sB = 0.5: r = dA ^ sB: out = out & "bin.Double0.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: sB = 16777216: r = dA + sB: out = out & "bin.Double0.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: sB = 16777216: r = dA - sB: out = out & "bin.Double0.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: sB = 16777216: r = dA * sB: out = out & "bin.Double0.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: sB = 16777216: r = dA / sB: out = out & "bin.Double0.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: sB = 16777216: r = dA \ sB: out = out & "bin.Double0.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: sB = 16777216: r = dA Mod sB: out = out & "bin.Double0.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: sB = 16777216: r = dA ^ sB: out = out & "bin.Double0.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: sB = 3.0E+38: r = dA + sB: out = out & "bin.Double0.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: sB = 3.0E+38: r = dA - sB: out = out & "bin.Double0.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: sB = 3.0E+38: r = dA * sB: out = out & "bin.Double0.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: sB = 3.0E+38: r = dA / sB: out = out & "bin.Double0.div.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: sB = 3.0E+38: r = dA \ sB: out = out & "bin.Double0.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: sB = 3.0E+38: r = dA Mod sB: out = out & "bin.Double0.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: sB = 3.0E+38: r = dA ^ sB: out = out & "bin.Double0.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: sB = 0.5: r = dA + sB: out = out & "bin.Double1.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: sB = 0.5: r = dA - sB: out = out & "bin.Double1.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: sB = 0.5: r = dA * sB: out = out & "bin.Double1.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: sB = 0.5: r = dA / sB: out = out & "bin.Double1.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: sB = 0.5: r = dA \ sB: out = out & "bin.Double1.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: sB = 0.5: r = dA Mod sB: out = out & "bin.Double1.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: sB = 0.5: r = dA ^ sB: out = out & "bin.Double1.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: sB = 16777216: r = dA + sB: out = out & "bin.Double1.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: sB = 16777216: r = dA - sB: out = out & "bin.Double1.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: sB = 16777216: r = dA * sB: out = out & "bin.Double1.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: sB = 16777216: r = dA / sB: out = out & "bin.Double1.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: sB = 16777216: r = dA \ sB: out = out & "bin.Double1.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: sB = 16777216: r = dA Mod sB: out = out & "bin.Double1.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: sB = 16777216: r = dA ^ sB: out = out & "bin.Double1.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: sB = 3.0E+38: r = dA + sB: out = out & "bin.Double1.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: sB = 3.0E+38: r = dA - sB: out = out & "bin.Double1.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: sB = 3.0E+38: r = dA * sB: out = out & "bin.Double1.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: sB = 3.0E+38: r = dA / sB: out = out & "bin.Double1.div.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: sB = 3.0E+38: r = dA \ sB: out = out & "bin.Double1.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: sB = 3.0E+38: r = dA Mod sB: out = out & "bin.Double1.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: sB = 3.0E+38: r = dA ^ sB: out = out & "bin.Double1.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: sB = 0.5: r = dA + sB: out = out & "bin.Double2.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: sB = 0.5: r = dA - sB: out = out & "bin.Double2.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: sB = 0.5: r = dA * sB: out = out & "bin.Double2.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: sB = 0.5: r = dA / sB: out = out & "bin.Double2.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: sB = 0.5: r = dA \ sB: out = out & "bin.Double2.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: sB = 0.5: r = dA Mod sB: out = out & "bin.Double2.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: sB = 0.5: r = dA ^ sB: out = out & "bin.Double2.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: sB = 16777216: r = dA + sB: out = out & "bin.Double2.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: sB = 16777216: r = dA - sB: out = out & "bin.Double2.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: sB = 16777216: r = dA * sB: out = out & "bin.Double2.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: sB = 16777216: r = dA / sB: out = out & "bin.Double2.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: sB = 16777216: r = dA \ sB: out = out & "bin.Double2.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: sB = 16777216: r = dA Mod sB: out = out & "bin.Double2.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: sB = 16777216: r = dA ^ sB: out = out & "bin.Double2.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: sB = 3.0E+38: r = dA + sB: out = out & "bin.Double2.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: sB = 3.0E+38: r = dA - sB: out = out & "bin.Double2.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: sB = 3.0E+38: r = dA * sB: out = out & "bin.Double2.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: sB = 3.0E+38: r = dA / sB: out = out & "bin.Double2.div.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: sB = 3.0E+38: r = dA \ sB: out = out & "bin.Double2.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: sB = 3.0E+38: r = dA Mod sB: out = out & "bin.Double2.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: sB = 3.0E+38: r = dA ^ sB: out = out & "bin.Double2.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: dB = 0.5: r = dA + dB: out = out & "bin.Double0.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: dB = 0.5: r = dA - dB: out = out & "bin.Double0.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: dB = 0.5: r = dA * dB: out = out & "bin.Double0.mul.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: dB = 0.5: r = dA / dB: out = out & "bin.Double0.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: dB = 0.5: r = dA \ dB: out = out & "bin.Double0.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: dB = 0.5: r = dA Mod dB: out = out & "bin.Double0.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: dB = 0.5: r = dA ^ dB: out = out & "bin.Double0.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: dB = 2.5: r = dA + dB: out = out & "bin.Double0.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: dB = 2.5: r = dA - dB: out = out & "bin.Double0.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: dB = 2.5: r = dA * dB: out = out & "bin.Double0.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: dB = 2.5: r = dA / dB: out = out & "bin.Double0.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: dB = 2.5: r = dA \ dB: out = out & "bin.Double0.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: dB = 2.5: r = dA Mod dB: out = out & "bin.Double0.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: dB = 2.5: r = dA ^ dB: out = out & "bin.Double0.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: dB = 1.0E+308: r = dA + dB: out = out & "bin.Double0.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: dB = 1.0E+308: r = dA - dB: out = out & "bin.Double0.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: dB = 1.0E+308: r = dA * dB: out = out & "bin.Double0.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: dB = 1.0E+308: r = dA / dB: out = out & "bin.Double0.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: dB = 1.0E+308: r = dA \ dB: out = out & "bin.Double0.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: dB = 1.0E+308: r = dA Mod dB: out = out & "bin.Double0.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: dB = 1.0E+308: r = dA ^ dB: out = out & "bin.Double0.pow.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: dB = 0.5: r = dA + dB: out = out & "bin.Double1.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: dB = 0.5: r = dA - dB: out = out & "bin.Double1.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: dB = 0.5: r = dA * dB: out = out & "bin.Double1.mul.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: dB = 0.5: r = dA / dB: out = out & "bin.Double1.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: dB = 0.5: r = dA \ dB: out = out & "bin.Double1.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: dB = 0.5: r = dA Mod dB: out = out & "bin.Double1.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: dB = 0.5: r = dA ^ dB: out = out & "bin.Double1.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: dB = 2.5: r = dA + dB: out = out & "bin.Double1.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: dB = 2.5: r = dA - dB: out = out & "bin.Double1.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: dB = 2.5: r = dA * dB: out = out & "bin.Double1.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: dB = 2.5: r = dA / dB: out = out & "bin.Double1.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: dB = 2.5: r = dA \ dB: out = out & "bin.Double1.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: dB = 2.5: r = dA Mod dB: out = out & "bin.Double1.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: dB = 2.5: r = dA ^ dB: out = out & "bin.Double1.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: dB = 1.0E+308: r = dA + dB: out = out & "bin.Double1.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: dB = 1.0E+308: r = dA - dB: out = out & "bin.Double1.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: dB = 1.0E+308: r = dA * dB: out = out & "bin.Double1.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: dB = 1.0E+308: r = dA / dB: out = out & "bin.Double1.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: dB = 1.0E+308: r = dA \ dB: out = out & "bin.Double1.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: dB = 1.0E+308: r = dA Mod dB: out = out & "bin.Double1.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: dB = 1.0E+308: r = dA ^ dB: out = out & "bin.Double1.pow.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: dB = 0.5: r = dA + dB: out = out & "bin.Double2.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: dB = 0.5: r = dA - dB: out = out & "bin.Double2.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: dB = 0.5: r = dA * dB: out = out & "bin.Double2.mul.Double0=" & Rz(Err.Number, r) & vbLf
  P023 = out
End Function
Function P024() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: dA = 1.0E+308: dB = 0.5: r = dA / dB: out = out & "bin.Double2.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: dB = 0.5: r = dA \ dB: out = out & "bin.Double2.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: dB = 0.5: r = dA Mod dB: out = out & "bin.Double2.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: dB = 0.5: r = dA ^ dB: out = out & "bin.Double2.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: dB = 2.5: r = dA + dB: out = out & "bin.Double2.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: dB = 2.5: r = dA - dB: out = out & "bin.Double2.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: dB = 2.5: r = dA * dB: out = out & "bin.Double2.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: dB = 2.5: r = dA / dB: out = out & "bin.Double2.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: dB = 2.5: r = dA \ dB: out = out & "bin.Double2.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: dB = 2.5: r = dA Mod dB: out = out & "bin.Double2.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: dB = 2.5: r = dA ^ dB: out = out & "bin.Double2.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: dB = 1.0E+308: r = dA + dB: out = out & "bin.Double2.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: dB = 1.0E+308: r = dA - dB: out = out & "bin.Double2.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: dB = 1.0E+308: r = dA * dB: out = out & "bin.Double2.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: dB = 1.0E+308: r = dA / dB: out = out & "bin.Double2.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: dB = 1.0E+308: r = dA \ dB: out = out & "bin.Double2.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: dB = 1.0E+308: r = dA Mod dB: out = out & "bin.Double2.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: dB = 1.0E+308: r = dA ^ dB: out = out & "bin.Double2.pow.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: cB = -0.5@: r = dA + cB: out = out & "bin.Double0.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: cB = -0.5@: r = dA - cB: out = out & "bin.Double0.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: cB = -0.5@: r = dA * cB: out = out & "bin.Double0.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: cB = -0.5@: r = dA / cB: out = out & "bin.Double0.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: cB = -0.5@: r = dA \ cB: out = out & "bin.Double0.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: cB = -0.5@: r = dA Mod cB: out = out & "bin.Double0.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: cB = -0.5@: r = dA ^ cB: out = out & "bin.Double0.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: cB = 1.5@: r = dA + cB: out = out & "bin.Double0.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: cB = 1.5@: r = dA - cB: out = out & "bin.Double0.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: cB = 1.5@: r = dA * cB: out = out & "bin.Double0.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: cB = 1.5@: r = dA / cB: out = out & "bin.Double0.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: cB = 1.5@: r = dA \ cB: out = out & "bin.Double0.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: cB = 1.5@: r = dA Mod cB: out = out & "bin.Double0.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: cB = 1.5@: r = dA ^ cB: out = out & "bin.Double0.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: cB = 922337203685477.5807@: r = dA + cB: out = out & "bin.Double0.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: cB = 922337203685477.5807@: r = dA - cB: out = out & "bin.Double0.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: cB = 922337203685477.5807@: r = dA * cB: out = out & "bin.Double0.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: cB = 922337203685477.5807@: r = dA / cB: out = out & "bin.Double0.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: cB = 922337203685477.5807@: r = dA \ cB: out = out & "bin.Double0.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: cB = 922337203685477.5807@: r = dA Mod cB: out = out & "bin.Double0.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: cB = 922337203685477.5807@: r = dA ^ cB: out = out & "bin.Double0.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: cB = -0.5@: r = dA + cB: out = out & "bin.Double1.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: cB = -0.5@: r = dA - cB: out = out & "bin.Double1.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: cB = -0.5@: r = dA * cB: out = out & "bin.Double1.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: cB = -0.5@: r = dA / cB: out = out & "bin.Double1.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: cB = -0.5@: r = dA \ cB: out = out & "bin.Double1.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: cB = -0.5@: r = dA Mod cB: out = out & "bin.Double1.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: cB = -0.5@: r = dA ^ cB: out = out & "bin.Double1.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: cB = 1.5@: r = dA + cB: out = out & "bin.Double1.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: cB = 1.5@: r = dA - cB: out = out & "bin.Double1.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: cB = 1.5@: r = dA * cB: out = out & "bin.Double1.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: cB = 1.5@: r = dA / cB: out = out & "bin.Double1.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: cB = 1.5@: r = dA \ cB: out = out & "bin.Double1.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: cB = 1.5@: r = dA Mod cB: out = out & "bin.Double1.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: cB = 1.5@: r = dA ^ cB: out = out & "bin.Double1.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: cB = 922337203685477.5807@: r = dA + cB: out = out & "bin.Double1.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: cB = 922337203685477.5807@: r = dA - cB: out = out & "bin.Double1.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: cB = 922337203685477.5807@: r = dA * cB: out = out & "bin.Double1.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: cB = 922337203685477.5807@: r = dA / cB: out = out & "bin.Double1.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: cB = 922337203685477.5807@: r = dA \ cB: out = out & "bin.Double1.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: cB = 922337203685477.5807@: r = dA Mod cB: out = out & "bin.Double1.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: cB = 922337203685477.5807@: r = dA ^ cB: out = out & "bin.Double1.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: cB = -0.5@: r = dA + cB: out = out & "bin.Double2.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: cB = -0.5@: r = dA - cB: out = out & "bin.Double2.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: cB = -0.5@: r = dA * cB: out = out & "bin.Double2.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: cB = -0.5@: r = dA / cB: out = out & "bin.Double2.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: cB = -0.5@: r = dA \ cB: out = out & "bin.Double2.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: cB = -0.5@: r = dA Mod cB: out = out & "bin.Double2.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: cB = -0.5@: r = dA ^ cB: out = out & "bin.Double2.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: cB = 1.5@: r = dA + cB: out = out & "bin.Double2.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: cB = 1.5@: r = dA - cB: out = out & "bin.Double2.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: cB = 1.5@: r = dA * cB: out = out & "bin.Double2.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: cB = 1.5@: r = dA / cB: out = out & "bin.Double2.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: cB = 1.5@: r = dA \ cB: out = out & "bin.Double2.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: cB = 1.5@: r = dA Mod cB: out = out & "bin.Double2.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: cB = 1.5@: r = dA ^ cB: out = out & "bin.Double2.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: cB = 922337203685477.5807@: r = dA + cB: out = out & "bin.Double2.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: cB = 922337203685477.5807@: r = dA - cB: out = out & "bin.Double2.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: cB = 922337203685477.5807@: r = dA * cB: out = out & "bin.Double2.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: cB = 922337203685477.5807@: r = dA / cB: out = out & "bin.Double2.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: cB = 922337203685477.5807@: r = dA \ cB: out = out & "bin.Double2.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: cB = 922337203685477.5807@: r = dA Mod cB: out = out & "bin.Double2.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: cB = 922337203685477.5807@: r = dA ^ cB: out = out & "bin.Double2.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: vB = Empty: r = dA + vB: out = out & "bin.Double0.add.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: vB = Empty: r = dA - vB: out = out & "bin.Double0.sub.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: vB = Empty: r = dA * vB: out = out & "bin.Double0.mul.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: vB = Empty: r = dA / vB: out = out & "bin.Double0.div.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: vB = Empty: r = dA \ vB: out = out & "bin.Double0.idiv.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: vB = Empty: r = dA Mod vB: out = out & "bin.Double0.mod.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: vB = Empty: r = dA ^ vB: out = out & "bin.Double0.pow.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: vB = 7: r = dA + vB: out = out & "bin.Double0.add.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: vB = 7: r = dA - vB: out = out & "bin.Double0.sub.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: vB = 7: r = dA * vB: out = out & "bin.Double0.mul.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: vB = 7: r = dA / vB: out = out & "bin.Double0.div.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: vB = 7: r = dA \ vB: out = out & "bin.Double0.idiv.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: vB = 7: r = dA Mod vB: out = out & "bin.Double0.mod.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: vB = 7: r = dA ^ vB: out = out & "bin.Double0.pow.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: vB = 2147483647: r = dA + vB: out = out & "bin.Double0.add.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: vB = 2147483647: r = dA - vB: out = out & "bin.Double0.sub.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: vB = 2147483647: r = dA * vB: out = out & "bin.Double0.mul.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: vB = 2147483647: r = dA / vB: out = out & "bin.Double0.div.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: vB = 2147483647: r = dA \ vB: out = out & "bin.Double0.idiv.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: vB = 2147483647: r = dA Mod vB: out = out & "bin.Double0.mod.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: vB = 2147483647: r = dA ^ vB: out = out & "bin.Double0.pow.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: vB = Empty: r = dA + vB: out = out & "bin.Double1.add.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: vB = Empty: r = dA - vB: out = out & "bin.Double1.sub.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: vB = Empty: r = dA * vB: out = out & "bin.Double1.mul.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: vB = Empty: r = dA / vB: out = out & "bin.Double1.div.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: vB = Empty: r = dA \ vB: out = out & "bin.Double1.idiv.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: vB = Empty: r = dA Mod vB: out = out & "bin.Double1.mod.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: vB = Empty: r = dA ^ vB: out = out & "bin.Double1.pow.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: vB = 7: r = dA + vB: out = out & "bin.Double1.add.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: vB = 7: r = dA - vB: out = out & "bin.Double1.sub.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: vB = 7: r = dA * vB: out = out & "bin.Double1.mul.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: vB = 7: r = dA / vB: out = out & "bin.Double1.div.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: vB = 7: r = dA \ vB: out = out & "bin.Double1.idiv.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: vB = 7: r = dA Mod vB: out = out & "bin.Double1.mod.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: vB = 7: r = dA ^ vB: out = out & "bin.Double1.pow.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: vB = 2147483647: r = dA + vB: out = out & "bin.Double1.add.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: vB = 2147483647: r = dA - vB: out = out & "bin.Double1.sub.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: vB = 2147483647: r = dA * vB: out = out & "bin.Double1.mul.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: vB = 2147483647: r = dA / vB: out = out & "bin.Double1.div.Variant2=" & Rz(Err.Number, r) & vbLf
  P024 = out
End Function
Function P025() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: dA = 2.5: vB = 2147483647: r = dA \ vB: out = out & "bin.Double1.idiv.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: vB = 2147483647: r = dA Mod vB: out = out & "bin.Double1.mod.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: vB = 2147483647: r = dA ^ vB: out = out & "bin.Double1.pow.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: vB = Empty: r = dA + vB: out = out & "bin.Double2.add.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: vB = Empty: r = dA - vB: out = out & "bin.Double2.sub.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: vB = Empty: r = dA * vB: out = out & "bin.Double2.mul.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: vB = Empty: r = dA / vB: out = out & "bin.Double2.div.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: vB = Empty: r = dA \ vB: out = out & "bin.Double2.idiv.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: vB = Empty: r = dA Mod vB: out = out & "bin.Double2.mod.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: vB = Empty: r = dA ^ vB: out = out & "bin.Double2.pow.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: vB = 7: r = dA + vB: out = out & "bin.Double2.add.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: vB = 7: r = dA - vB: out = out & "bin.Double2.sub.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: vB = 7: r = dA * vB: out = out & "bin.Double2.mul.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: vB = 7: r = dA / vB: out = out & "bin.Double2.div.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: vB = 7: r = dA \ vB: out = out & "bin.Double2.idiv.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: vB = 7: r = dA Mod vB: out = out & "bin.Double2.mod.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: vB = 7: r = dA ^ vB: out = out & "bin.Double2.pow.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: vB = 2147483647: r = dA + vB: out = out & "bin.Double2.add.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: vB = 2147483647: r = dA - vB: out = out & "bin.Double2.sub.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: vB = 2147483647: r = dA * vB: out = out & "bin.Double2.mul.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: vB = 2147483647: r = dA / vB: out = out & "bin.Double2.div.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: vB = 2147483647: r = dA \ vB: out = out & "bin.Double2.idiv.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: vB = 2147483647: r = dA Mod vB: out = out & "bin.Double2.mod.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: vB = 2147483647: r = dA ^ vB: out = out & "bin.Double2.pow.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: bB = 0: r = cA + bB: out = out & "bin.Currency0.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: bB = 0: r = cA - bB: out = out & "bin.Currency0.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: bB = 0: r = cA * bB: out = out & "bin.Currency0.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: bB = 0: r = cA / bB: out = out & "bin.Currency0.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: bB = 0: r = cA \ bB: out = out & "bin.Currency0.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: bB = 0: r = cA Mod bB: out = out & "bin.Currency0.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: bB = 0: r = cA ^ bB: out = out & "bin.Currency0.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: bB = 200: r = cA + bB: out = out & "bin.Currency0.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: bB = 200: r = cA - bB: out = out & "bin.Currency0.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: bB = 200: r = cA * bB: out = out & "bin.Currency0.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: bB = 200: r = cA / bB: out = out & "bin.Currency0.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: bB = 200: r = cA \ bB: out = out & "bin.Currency0.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: bB = 200: r = cA Mod bB: out = out & "bin.Currency0.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: bB = 200: r = cA ^ bB: out = out & "bin.Currency0.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: bB = 255: r = cA + bB: out = out & "bin.Currency0.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: bB = 255: r = cA - bB: out = out & "bin.Currency0.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: bB = 255: r = cA * bB: out = out & "bin.Currency0.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: bB = 255: r = cA / bB: out = out & "bin.Currency0.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: bB = 255: r = cA \ bB: out = out & "bin.Currency0.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: bB = 255: r = cA Mod bB: out = out & "bin.Currency0.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: bB = 255: r = cA ^ bB: out = out & "bin.Currency0.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: bB = 0: r = cA + bB: out = out & "bin.Currency1.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: bB = 0: r = cA - bB: out = out & "bin.Currency1.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: bB = 0: r = cA * bB: out = out & "bin.Currency1.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: bB = 0: r = cA / bB: out = out & "bin.Currency1.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: bB = 0: r = cA \ bB: out = out & "bin.Currency1.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: bB = 0: r = cA Mod bB: out = out & "bin.Currency1.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: bB = 0: r = cA ^ bB: out = out & "bin.Currency1.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: bB = 200: r = cA + bB: out = out & "bin.Currency1.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: bB = 200: r = cA - bB: out = out & "bin.Currency1.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: bB = 200: r = cA * bB: out = out & "bin.Currency1.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: bB = 200: r = cA / bB: out = out & "bin.Currency1.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: bB = 200: r = cA \ bB: out = out & "bin.Currency1.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: bB = 200: r = cA Mod bB: out = out & "bin.Currency1.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: bB = 200: r = cA ^ bB: out = out & "bin.Currency1.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: bB = 255: r = cA + bB: out = out & "bin.Currency1.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: bB = 255: r = cA - bB: out = out & "bin.Currency1.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: bB = 255: r = cA * bB: out = out & "bin.Currency1.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: bB = 255: r = cA / bB: out = out & "bin.Currency1.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: bB = 255: r = cA \ bB: out = out & "bin.Currency1.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: bB = 255: r = cA Mod bB: out = out & "bin.Currency1.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: bB = 255: r = cA ^ bB: out = out & "bin.Currency1.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bB = 0: r = cA + bB: out = out & "bin.Currency2.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bB = 0: r = cA - bB: out = out & "bin.Currency2.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bB = 0: r = cA * bB: out = out & "bin.Currency2.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bB = 0: r = cA / bB: out = out & "bin.Currency2.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bB = 0: r = cA \ bB: out = out & "bin.Currency2.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bB = 0: r = cA Mod bB: out = out & "bin.Currency2.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bB = 0: r = cA ^ bB: out = out & "bin.Currency2.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bB = 200: r = cA + bB: out = out & "bin.Currency2.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bB = 200: r = cA - bB: out = out & "bin.Currency2.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bB = 200: r = cA * bB: out = out & "bin.Currency2.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bB = 200: r = cA / bB: out = out & "bin.Currency2.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bB = 200: r = cA \ bB: out = out & "bin.Currency2.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bB = 200: r = cA Mod bB: out = out & "bin.Currency2.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bB = 200: r = cA ^ bB: out = out & "bin.Currency2.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bB = 255: r = cA + bB: out = out & "bin.Currency2.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bB = 255: r = cA - bB: out = out & "bin.Currency2.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bB = 255: r = cA * bB: out = out & "bin.Currency2.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bB = 255: r = cA / bB: out = out & "bin.Currency2.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bB = 255: r = cA \ bB: out = out & "bin.Currency2.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bB = 255: r = cA Mod bB: out = out & "bin.Currency2.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bB = 255: r = cA ^ bB: out = out & "bin.Currency2.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: iB = -32768: r = cA + iB: out = out & "bin.Currency0.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: iB = -32768: r = cA - iB: out = out & "bin.Currency0.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: iB = -32768: r = cA * iB: out = out & "bin.Currency0.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: iB = -32768: r = cA / iB: out = out & "bin.Currency0.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: iB = -32768: r = cA \ iB: out = out & "bin.Currency0.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: iB = -32768: r = cA Mod iB: out = out & "bin.Currency0.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: iB = -32768: r = cA ^ iB: out = out & "bin.Currency0.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: iB = 1: r = cA + iB: out = out & "bin.Currency0.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: iB = 1: r = cA - iB: out = out & "bin.Currency0.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: iB = 1: r = cA * iB: out = out & "bin.Currency0.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: iB = 1: r = cA / iB: out = out & "bin.Currency0.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: iB = 1: r = cA \ iB: out = out & "bin.Currency0.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: iB = 1: r = cA Mod iB: out = out & "bin.Currency0.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: iB = 1: r = cA ^ iB: out = out & "bin.Currency0.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: iB = 32767: r = cA + iB: out = out & "bin.Currency0.add.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: iB = 32767: r = cA - iB: out = out & "bin.Currency0.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: iB = 32767: r = cA * iB: out = out & "bin.Currency0.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: iB = 32767: r = cA / iB: out = out & "bin.Currency0.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: iB = 32767: r = cA \ iB: out = out & "bin.Currency0.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: iB = 32767: r = cA Mod iB: out = out & "bin.Currency0.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: iB = 32767: r = cA ^ iB: out = out & "bin.Currency0.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: iB = -32768: r = cA + iB: out = out & "bin.Currency1.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: iB = -32768: r = cA - iB: out = out & "bin.Currency1.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: iB = -32768: r = cA * iB: out = out & "bin.Currency1.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: iB = -32768: r = cA / iB: out = out & "bin.Currency1.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: iB = -32768: r = cA \ iB: out = out & "bin.Currency1.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: iB = -32768: r = cA Mod iB: out = out & "bin.Currency1.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: iB = -32768: r = cA ^ iB: out = out & "bin.Currency1.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: iB = 1: r = cA + iB: out = out & "bin.Currency1.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: iB = 1: r = cA - iB: out = out & "bin.Currency1.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: iB = 1: r = cA * iB: out = out & "bin.Currency1.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: iB = 1: r = cA / iB: out = out & "bin.Currency1.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: iB = 1: r = cA \ iB: out = out & "bin.Currency1.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  P025 = out
End Function
Function P026() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: cA = 1.5@: iB = 1: r = cA Mod iB: out = out & "bin.Currency1.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: iB = 1: r = cA ^ iB: out = out & "bin.Currency1.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: iB = 32767: r = cA + iB: out = out & "bin.Currency1.add.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: iB = 32767: r = cA - iB: out = out & "bin.Currency1.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: iB = 32767: r = cA * iB: out = out & "bin.Currency1.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: iB = 32767: r = cA / iB: out = out & "bin.Currency1.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: iB = 32767: r = cA \ iB: out = out & "bin.Currency1.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: iB = 32767: r = cA Mod iB: out = out & "bin.Currency1.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: iB = 32767: r = cA ^ iB: out = out & "bin.Currency1.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iB = -32768: r = cA + iB: out = out & "bin.Currency2.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iB = -32768: r = cA - iB: out = out & "bin.Currency2.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iB = -32768: r = cA * iB: out = out & "bin.Currency2.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iB = -32768: r = cA / iB: out = out & "bin.Currency2.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iB = -32768: r = cA \ iB: out = out & "bin.Currency2.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iB = -32768: r = cA Mod iB: out = out & "bin.Currency2.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iB = -32768: r = cA ^ iB: out = out & "bin.Currency2.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iB = 1: r = cA + iB: out = out & "bin.Currency2.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iB = 1: r = cA - iB: out = out & "bin.Currency2.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iB = 1: r = cA * iB: out = out & "bin.Currency2.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iB = 1: r = cA / iB: out = out & "bin.Currency2.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iB = 1: r = cA \ iB: out = out & "bin.Currency2.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iB = 1: r = cA Mod iB: out = out & "bin.Currency2.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iB = 1: r = cA ^ iB: out = out & "bin.Currency2.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iB = 32767: r = cA + iB: out = out & "bin.Currency2.add.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iB = 32767: r = cA - iB: out = out & "bin.Currency2.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iB = 32767: r = cA * iB: out = out & "bin.Currency2.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iB = 32767: r = cA / iB: out = out & "bin.Currency2.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iB = 32767: r = cA \ iB: out = out & "bin.Currency2.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iB = 32767: r = cA Mod iB: out = out & "bin.Currency2.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iB = 32767: r = cA ^ iB: out = out & "bin.Currency2.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: lB = -2147483648: r = cA + lB: out = out & "bin.Currency0.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: lB = -2147483648: r = cA - lB: out = out & "bin.Currency0.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: lB = -2147483648: r = cA * lB: out = out & "bin.Currency0.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: lB = -2147483648: r = cA / lB: out = out & "bin.Currency0.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: lB = -2147483648: r = cA \ lB: out = out & "bin.Currency0.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: lB = -2147483648: r = cA Mod lB: out = out & "bin.Currency0.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: lB = -2147483648: r = cA ^ lB: out = out & "bin.Currency0.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: lB = 65536: r = cA + lB: out = out & "bin.Currency0.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: lB = 65536: r = cA - lB: out = out & "bin.Currency0.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: lB = 65536: r = cA * lB: out = out & "bin.Currency0.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: lB = 65536: r = cA / lB: out = out & "bin.Currency0.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: lB = 65536: r = cA \ lB: out = out & "bin.Currency0.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: lB = 65536: r = cA Mod lB: out = out & "bin.Currency0.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: lB = 65536: r = cA ^ lB: out = out & "bin.Currency0.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: lB = 2147483647: r = cA + lB: out = out & "bin.Currency0.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: lB = 2147483647: r = cA - lB: out = out & "bin.Currency0.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: lB = 2147483647: r = cA * lB: out = out & "bin.Currency0.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: lB = 2147483647: r = cA / lB: out = out & "bin.Currency0.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: lB = 2147483647: r = cA \ lB: out = out & "bin.Currency0.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: lB = 2147483647: r = cA Mod lB: out = out & "bin.Currency0.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: lB = 2147483647: r = cA ^ lB: out = out & "bin.Currency0.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: lB = -2147483648: r = cA + lB: out = out & "bin.Currency1.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: lB = -2147483648: r = cA - lB: out = out & "bin.Currency1.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: lB = -2147483648: r = cA * lB: out = out & "bin.Currency1.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: lB = -2147483648: r = cA / lB: out = out & "bin.Currency1.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: lB = -2147483648: r = cA \ lB: out = out & "bin.Currency1.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: lB = -2147483648: r = cA Mod lB: out = out & "bin.Currency1.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: lB = -2147483648: r = cA ^ lB: out = out & "bin.Currency1.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: lB = 65536: r = cA + lB: out = out & "bin.Currency1.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: lB = 65536: r = cA - lB: out = out & "bin.Currency1.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: lB = 65536: r = cA * lB: out = out & "bin.Currency1.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: lB = 65536: r = cA / lB: out = out & "bin.Currency1.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: lB = 65536: r = cA \ lB: out = out & "bin.Currency1.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: lB = 65536: r = cA Mod lB: out = out & "bin.Currency1.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: lB = 65536: r = cA ^ lB: out = out & "bin.Currency1.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: lB = 2147483647: r = cA + lB: out = out & "bin.Currency1.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: lB = 2147483647: r = cA - lB: out = out & "bin.Currency1.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: lB = 2147483647: r = cA * lB: out = out & "bin.Currency1.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: lB = 2147483647: r = cA / lB: out = out & "bin.Currency1.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: lB = 2147483647: r = cA \ lB: out = out & "bin.Currency1.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: lB = 2147483647: r = cA Mod lB: out = out & "bin.Currency1.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: lB = 2147483647: r = cA ^ lB: out = out & "bin.Currency1.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lB = -2147483648: r = cA + lB: out = out & "bin.Currency2.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lB = -2147483648: r = cA - lB: out = out & "bin.Currency2.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lB = -2147483648: r = cA * lB: out = out & "bin.Currency2.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lB = -2147483648: r = cA / lB: out = out & "bin.Currency2.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lB = -2147483648: r = cA \ lB: out = out & "bin.Currency2.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lB = -2147483648: r = cA Mod lB: out = out & "bin.Currency2.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lB = -2147483648: r = cA ^ lB: out = out & "bin.Currency2.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lB = 65536: r = cA + lB: out = out & "bin.Currency2.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lB = 65536: r = cA - lB: out = out & "bin.Currency2.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lB = 65536: r = cA * lB: out = out & "bin.Currency2.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lB = 65536: r = cA / lB: out = out & "bin.Currency2.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lB = 65536: r = cA \ lB: out = out & "bin.Currency2.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lB = 65536: r = cA Mod lB: out = out & "bin.Currency2.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lB = 65536: r = cA ^ lB: out = out & "bin.Currency2.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lB = 2147483647: r = cA + lB: out = out & "bin.Currency2.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lB = 2147483647: r = cA - lB: out = out & "bin.Currency2.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lB = 2147483647: r = cA * lB: out = out & "bin.Currency2.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lB = 2147483647: r = cA / lB: out = out & "bin.Currency2.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lB = 2147483647: r = cA \ lB: out = out & "bin.Currency2.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lB = 2147483647: r = cA Mod lB: out = out & "bin.Currency2.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lB = 2147483647: r = cA ^ lB: out = out & "bin.Currency2.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: qB = -9223372036854775807^: r = cA + qB: out = out & "bin.Currency0.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: qB = -9223372036854775807^: r = cA - qB: out = out & "bin.Currency0.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: qB = -9223372036854775807^: r = cA * qB: out = out & "bin.Currency0.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: qB = -9223372036854775807^: r = cA / qB: out = out & "bin.Currency0.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: qB = -9223372036854775807^: r = cA \ qB: out = out & "bin.Currency0.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: qB = -9223372036854775807^: r = cA Mod qB: out = out & "bin.Currency0.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: qB = -9223372036854775807^: r = cA ^ qB: out = out & "bin.Currency0.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: qB = 3: r = cA + qB: out = out & "bin.Currency0.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: qB = 3: r = cA - qB: out = out & "bin.Currency0.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: qB = 3: r = cA * qB: out = out & "bin.Currency0.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: qB = 3: r = cA / qB: out = out & "bin.Currency0.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: qB = 3: r = cA \ qB: out = out & "bin.Currency0.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: qB = 3: r = cA Mod qB: out = out & "bin.Currency0.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: qB = 3: r = cA ^ qB: out = out & "bin.Currency0.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: qB = 9223372036854775807^: r = cA + qB: out = out & "bin.Currency0.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: qB = 9223372036854775807^: r = cA - qB: out = out & "bin.Currency0.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: qB = 9223372036854775807^: r = cA * qB: out = out & "bin.Currency0.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: qB = 9223372036854775807^: r = cA / qB: out = out & "bin.Currency0.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: qB = 9223372036854775807^: r = cA \ qB: out = out & "bin.Currency0.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: qB = 9223372036854775807^: r = cA Mod qB: out = out & "bin.Currency0.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: qB = 9223372036854775807^: r = cA ^ qB: out = out & "bin.Currency0.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: qB = -9223372036854775807^: r = cA + qB: out = out & "bin.Currency1.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: qB = -9223372036854775807^: r = cA - qB: out = out & "bin.Currency1.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: qB = -9223372036854775807^: r = cA * qB: out = out & "bin.Currency1.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: qB = -9223372036854775807^: r = cA / qB: out = out & "bin.Currency1.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: qB = -9223372036854775807^: r = cA \ qB: out = out & "bin.Currency1.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: qB = -9223372036854775807^: r = cA Mod qB: out = out & "bin.Currency1.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  P026 = out
End Function
Function P027() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: cA = 1.5@: qB = -9223372036854775807^: r = cA ^ qB: out = out & "bin.Currency1.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: qB = 3: r = cA + qB: out = out & "bin.Currency1.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: qB = 3: r = cA - qB: out = out & "bin.Currency1.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: qB = 3: r = cA * qB: out = out & "bin.Currency1.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: qB = 3: r = cA / qB: out = out & "bin.Currency1.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: qB = 3: r = cA \ qB: out = out & "bin.Currency1.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: qB = 3: r = cA Mod qB: out = out & "bin.Currency1.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: qB = 3: r = cA ^ qB: out = out & "bin.Currency1.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: qB = 9223372036854775807^: r = cA + qB: out = out & "bin.Currency1.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: qB = 9223372036854775807^: r = cA - qB: out = out & "bin.Currency1.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: qB = 9223372036854775807^: r = cA * qB: out = out & "bin.Currency1.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: qB = 9223372036854775807^: r = cA / qB: out = out & "bin.Currency1.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: qB = 9223372036854775807^: r = cA \ qB: out = out & "bin.Currency1.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: qB = 9223372036854775807^: r = cA Mod qB: out = out & "bin.Currency1.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: qB = 9223372036854775807^: r = cA ^ qB: out = out & "bin.Currency1.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qB = -9223372036854775807^: r = cA + qB: out = out & "bin.Currency2.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qB = -9223372036854775807^: r = cA - qB: out = out & "bin.Currency2.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qB = -9223372036854775807^: r = cA * qB: out = out & "bin.Currency2.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qB = -9223372036854775807^: r = cA / qB: out = out & "bin.Currency2.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qB = -9223372036854775807^: r = cA \ qB: out = out & "bin.Currency2.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qB = -9223372036854775807^: r = cA Mod qB: out = out & "bin.Currency2.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qB = -9223372036854775807^: r = cA ^ qB: out = out & "bin.Currency2.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qB = 3: r = cA + qB: out = out & "bin.Currency2.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qB = 3: r = cA - qB: out = out & "bin.Currency2.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qB = 3: r = cA * qB: out = out & "bin.Currency2.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qB = 3: r = cA / qB: out = out & "bin.Currency2.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qB = 3: r = cA \ qB: out = out & "bin.Currency2.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qB = 3: r = cA Mod qB: out = out & "bin.Currency2.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qB = 3: r = cA ^ qB: out = out & "bin.Currency2.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qB = 9223372036854775807^: r = cA + qB: out = out & "bin.Currency2.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qB = 9223372036854775807^: r = cA - qB: out = out & "bin.Currency2.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qB = 9223372036854775807^: r = cA * qB: out = out & "bin.Currency2.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qB = 9223372036854775807^: r = cA / qB: out = out & "bin.Currency2.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qB = 9223372036854775807^: r = cA \ qB: out = out & "bin.Currency2.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qB = 9223372036854775807^: r = cA Mod qB: out = out & "bin.Currency2.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qB = 9223372036854775807^: r = cA ^ qB: out = out & "bin.Currency2.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: sB = 0.5: r = cA + sB: out = out & "bin.Currency0.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: sB = 0.5: r = cA - sB: out = out & "bin.Currency0.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: sB = 0.5: r = cA * sB: out = out & "bin.Currency0.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: sB = 0.5: r = cA / sB: out = out & "bin.Currency0.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: sB = 0.5: r = cA \ sB: out = out & "bin.Currency0.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: sB = 0.5: r = cA Mod sB: out = out & "bin.Currency0.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: sB = 0.5: r = cA ^ sB: out = out & "bin.Currency0.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: sB = 16777216: r = cA + sB: out = out & "bin.Currency0.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: sB = 16777216: r = cA - sB: out = out & "bin.Currency0.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: sB = 16777216: r = cA * sB: out = out & "bin.Currency0.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: sB = 16777216: r = cA / sB: out = out & "bin.Currency0.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: sB = 16777216: r = cA \ sB: out = out & "bin.Currency0.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: sB = 16777216: r = cA Mod sB: out = out & "bin.Currency0.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: sB = 16777216: r = cA ^ sB: out = out & "bin.Currency0.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: sB = 3.0E+38: r = cA + sB: out = out & "bin.Currency0.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: sB = 3.0E+38: r = cA - sB: out = out & "bin.Currency0.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: sB = 3.0E+38: r = cA * sB: out = out & "bin.Currency0.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: sB = 3.0E+38: r = cA / sB: out = out & "bin.Currency0.div.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: sB = 3.0E+38: r = cA \ sB: out = out & "bin.Currency0.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: sB = 3.0E+38: r = cA Mod sB: out = out & "bin.Currency0.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: sB = 3.0E+38: r = cA ^ sB: out = out & "bin.Currency0.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: sB = 0.5: r = cA + sB: out = out & "bin.Currency1.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: sB = 0.5: r = cA - sB: out = out & "bin.Currency1.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: sB = 0.5: r = cA * sB: out = out & "bin.Currency1.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: sB = 0.5: r = cA / sB: out = out & "bin.Currency1.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: sB = 0.5: r = cA \ sB: out = out & "bin.Currency1.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: sB = 0.5: r = cA Mod sB: out = out & "bin.Currency1.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: sB = 0.5: r = cA ^ sB: out = out & "bin.Currency1.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: sB = 16777216: r = cA + sB: out = out & "bin.Currency1.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: sB = 16777216: r = cA - sB: out = out & "bin.Currency1.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: sB = 16777216: r = cA * sB: out = out & "bin.Currency1.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: sB = 16777216: r = cA / sB: out = out & "bin.Currency1.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: sB = 16777216: r = cA \ sB: out = out & "bin.Currency1.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: sB = 16777216: r = cA Mod sB: out = out & "bin.Currency1.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: sB = 16777216: r = cA ^ sB: out = out & "bin.Currency1.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: sB = 3.0E+38: r = cA + sB: out = out & "bin.Currency1.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: sB = 3.0E+38: r = cA - sB: out = out & "bin.Currency1.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: sB = 3.0E+38: r = cA * sB: out = out & "bin.Currency1.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: sB = 3.0E+38: r = cA / sB: out = out & "bin.Currency1.div.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: sB = 3.0E+38: r = cA \ sB: out = out & "bin.Currency1.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: sB = 3.0E+38: r = cA Mod sB: out = out & "bin.Currency1.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: sB = 3.0E+38: r = cA ^ sB: out = out & "bin.Currency1.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sB = 0.5: r = cA + sB: out = out & "bin.Currency2.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sB = 0.5: r = cA - sB: out = out & "bin.Currency2.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sB = 0.5: r = cA * sB: out = out & "bin.Currency2.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sB = 0.5: r = cA / sB: out = out & "bin.Currency2.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sB = 0.5: r = cA \ sB: out = out & "bin.Currency2.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sB = 0.5: r = cA Mod sB: out = out & "bin.Currency2.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sB = 0.5: r = cA ^ sB: out = out & "bin.Currency2.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sB = 16777216: r = cA + sB: out = out & "bin.Currency2.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sB = 16777216: r = cA - sB: out = out & "bin.Currency2.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sB = 16777216: r = cA * sB: out = out & "bin.Currency2.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sB = 16777216: r = cA / sB: out = out & "bin.Currency2.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sB = 16777216: r = cA \ sB: out = out & "bin.Currency2.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sB = 16777216: r = cA Mod sB: out = out & "bin.Currency2.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sB = 16777216: r = cA ^ sB: out = out & "bin.Currency2.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sB = 3.0E+38: r = cA + sB: out = out & "bin.Currency2.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sB = 3.0E+38: r = cA - sB: out = out & "bin.Currency2.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sB = 3.0E+38: r = cA * sB: out = out & "bin.Currency2.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sB = 3.0E+38: r = cA / sB: out = out & "bin.Currency2.div.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sB = 3.0E+38: r = cA \ sB: out = out & "bin.Currency2.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sB = 3.0E+38: r = cA Mod sB: out = out & "bin.Currency2.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sB = 3.0E+38: r = cA ^ sB: out = out & "bin.Currency2.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: dB = 0.5: r = cA + dB: out = out & "bin.Currency0.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: dB = 0.5: r = cA - dB: out = out & "bin.Currency0.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: dB = 0.5: r = cA * dB: out = out & "bin.Currency0.mul.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: dB = 0.5: r = cA / dB: out = out & "bin.Currency0.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: dB = 0.5: r = cA \ dB: out = out & "bin.Currency0.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: dB = 0.5: r = cA Mod dB: out = out & "bin.Currency0.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: dB = 0.5: r = cA ^ dB: out = out & "bin.Currency0.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: dB = 2.5: r = cA + dB: out = out & "bin.Currency0.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: dB = 2.5: r = cA - dB: out = out & "bin.Currency0.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: dB = 2.5: r = cA * dB: out = out & "bin.Currency0.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: dB = 2.5: r = cA / dB: out = out & "bin.Currency0.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: dB = 2.5: r = cA \ dB: out = out & "bin.Currency0.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: dB = 2.5: r = cA Mod dB: out = out & "bin.Currency0.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: dB = 2.5: r = cA ^ dB: out = out & "bin.Currency0.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: dB = 1.0E+308: r = cA + dB: out = out & "bin.Currency0.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: dB = 1.0E+308: r = cA - dB: out = out & "bin.Currency0.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: dB = 1.0E+308: r = cA * dB: out = out & "bin.Currency0.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: dB = 1.0E+308: r = cA / dB: out = out & "bin.Currency0.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: dB = 1.0E+308: r = cA \ dB: out = out & "bin.Currency0.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: dB = 1.0E+308: r = cA Mod dB: out = out & "bin.Currency0.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: dB = 1.0E+308: r = cA ^ dB: out = out & "bin.Currency0.pow.Double2=" & Rz(Err.Number, r) & vbLf
  P027 = out
End Function
Function P028() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: cA = 1.5@: dB = 0.5: r = cA + dB: out = out & "bin.Currency1.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: dB = 0.5: r = cA - dB: out = out & "bin.Currency1.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: dB = 0.5: r = cA * dB: out = out & "bin.Currency1.mul.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: dB = 0.5: r = cA / dB: out = out & "bin.Currency1.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: dB = 0.5: r = cA \ dB: out = out & "bin.Currency1.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: dB = 0.5: r = cA Mod dB: out = out & "bin.Currency1.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: dB = 0.5: r = cA ^ dB: out = out & "bin.Currency1.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: dB = 2.5: r = cA + dB: out = out & "bin.Currency1.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: dB = 2.5: r = cA - dB: out = out & "bin.Currency1.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: dB = 2.5: r = cA * dB: out = out & "bin.Currency1.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: dB = 2.5: r = cA / dB: out = out & "bin.Currency1.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: dB = 2.5: r = cA \ dB: out = out & "bin.Currency1.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: dB = 2.5: r = cA Mod dB: out = out & "bin.Currency1.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: dB = 2.5: r = cA ^ dB: out = out & "bin.Currency1.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: dB = 1.0E+308: r = cA + dB: out = out & "bin.Currency1.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: dB = 1.0E+308: r = cA - dB: out = out & "bin.Currency1.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: dB = 1.0E+308: r = cA * dB: out = out & "bin.Currency1.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: dB = 1.0E+308: r = cA / dB: out = out & "bin.Currency1.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: dB = 1.0E+308: r = cA \ dB: out = out & "bin.Currency1.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: dB = 1.0E+308: r = cA Mod dB: out = out & "bin.Currency1.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: dB = 1.0E+308: r = cA ^ dB: out = out & "bin.Currency1.pow.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dB = 0.5: r = cA + dB: out = out & "bin.Currency2.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dB = 0.5: r = cA - dB: out = out & "bin.Currency2.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dB = 0.5: r = cA * dB: out = out & "bin.Currency2.mul.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dB = 0.5: r = cA / dB: out = out & "bin.Currency2.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dB = 0.5: r = cA \ dB: out = out & "bin.Currency2.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dB = 0.5: r = cA Mod dB: out = out & "bin.Currency2.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dB = 0.5: r = cA ^ dB: out = out & "bin.Currency2.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dB = 2.5: r = cA + dB: out = out & "bin.Currency2.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dB = 2.5: r = cA - dB: out = out & "bin.Currency2.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dB = 2.5: r = cA * dB: out = out & "bin.Currency2.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dB = 2.5: r = cA / dB: out = out & "bin.Currency2.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dB = 2.5: r = cA \ dB: out = out & "bin.Currency2.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dB = 2.5: r = cA Mod dB: out = out & "bin.Currency2.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dB = 2.5: r = cA ^ dB: out = out & "bin.Currency2.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dB = 1.0E+308: r = cA + dB: out = out & "bin.Currency2.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dB = 1.0E+308: r = cA - dB: out = out & "bin.Currency2.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dB = 1.0E+308: r = cA * dB: out = out & "bin.Currency2.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dB = 1.0E+308: r = cA / dB: out = out & "bin.Currency2.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dB = 1.0E+308: r = cA \ dB: out = out & "bin.Currency2.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dB = 1.0E+308: r = cA Mod dB: out = out & "bin.Currency2.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dB = 1.0E+308: r = cA ^ dB: out = out & "bin.Currency2.pow.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: cB = -0.5@: r = cA + cB: out = out & "bin.Currency0.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: cB = -0.5@: r = cA - cB: out = out & "bin.Currency0.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: cB = -0.5@: r = cA * cB: out = out & "bin.Currency0.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: cB = -0.5@: r = cA / cB: out = out & "bin.Currency0.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: cB = -0.5@: r = cA \ cB: out = out & "bin.Currency0.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: cB = -0.5@: r = cA Mod cB: out = out & "bin.Currency0.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: cB = -0.5@: r = cA ^ cB: out = out & "bin.Currency0.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: cB = 1.5@: r = cA + cB: out = out & "bin.Currency0.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: cB = 1.5@: r = cA - cB: out = out & "bin.Currency0.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: cB = 1.5@: r = cA * cB: out = out & "bin.Currency0.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: cB = 1.5@: r = cA / cB: out = out & "bin.Currency0.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: cB = 1.5@: r = cA \ cB: out = out & "bin.Currency0.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: cB = 1.5@: r = cA Mod cB: out = out & "bin.Currency0.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: cB = 1.5@: r = cA ^ cB: out = out & "bin.Currency0.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: cB = 922337203685477.5807@: r = cA + cB: out = out & "bin.Currency0.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: cB = 922337203685477.5807@: r = cA - cB: out = out & "bin.Currency0.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: cB = 922337203685477.5807@: r = cA * cB: out = out & "bin.Currency0.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: cB = 922337203685477.5807@: r = cA / cB: out = out & "bin.Currency0.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: cB = 922337203685477.5807@: r = cA \ cB: out = out & "bin.Currency0.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: cB = 922337203685477.5807@: r = cA Mod cB: out = out & "bin.Currency0.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: cB = 922337203685477.5807@: r = cA ^ cB: out = out & "bin.Currency0.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: cB = -0.5@: r = cA + cB: out = out & "bin.Currency1.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: cB = -0.5@: r = cA - cB: out = out & "bin.Currency1.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: cB = -0.5@: r = cA * cB: out = out & "bin.Currency1.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: cB = -0.5@: r = cA / cB: out = out & "bin.Currency1.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: cB = -0.5@: r = cA \ cB: out = out & "bin.Currency1.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: cB = -0.5@: r = cA Mod cB: out = out & "bin.Currency1.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: cB = -0.5@: r = cA ^ cB: out = out & "bin.Currency1.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: cB = 1.5@: r = cA + cB: out = out & "bin.Currency1.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: cB = 1.5@: r = cA - cB: out = out & "bin.Currency1.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: cB = 1.5@: r = cA * cB: out = out & "bin.Currency1.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: cB = 1.5@: r = cA / cB: out = out & "bin.Currency1.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: cB = 1.5@: r = cA \ cB: out = out & "bin.Currency1.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: cB = 1.5@: r = cA Mod cB: out = out & "bin.Currency1.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: cB = 1.5@: r = cA ^ cB: out = out & "bin.Currency1.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: cB = 922337203685477.5807@: r = cA + cB: out = out & "bin.Currency1.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: cB = 922337203685477.5807@: r = cA - cB: out = out & "bin.Currency1.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: cB = 922337203685477.5807@: r = cA * cB: out = out & "bin.Currency1.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: cB = 922337203685477.5807@: r = cA / cB: out = out & "bin.Currency1.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: cB = 922337203685477.5807@: r = cA \ cB: out = out & "bin.Currency1.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: cB = 922337203685477.5807@: r = cA Mod cB: out = out & "bin.Currency1.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: cB = 922337203685477.5807@: r = cA ^ cB: out = out & "bin.Currency1.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cB = -0.5@: r = cA + cB: out = out & "bin.Currency2.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cB = -0.5@: r = cA - cB: out = out & "bin.Currency2.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cB = -0.5@: r = cA * cB: out = out & "bin.Currency2.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cB = -0.5@: r = cA / cB: out = out & "bin.Currency2.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cB = -0.5@: r = cA \ cB: out = out & "bin.Currency2.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cB = -0.5@: r = cA Mod cB: out = out & "bin.Currency2.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cB = -0.5@: r = cA ^ cB: out = out & "bin.Currency2.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cB = 1.5@: r = cA + cB: out = out & "bin.Currency2.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cB = 1.5@: r = cA - cB: out = out & "bin.Currency2.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cB = 1.5@: r = cA * cB: out = out & "bin.Currency2.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cB = 1.5@: r = cA / cB: out = out & "bin.Currency2.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cB = 1.5@: r = cA \ cB: out = out & "bin.Currency2.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cB = 1.5@: r = cA Mod cB: out = out & "bin.Currency2.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cB = 1.5@: r = cA ^ cB: out = out & "bin.Currency2.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cB = 922337203685477.5807@: r = cA + cB: out = out & "bin.Currency2.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cB = 922337203685477.5807@: r = cA - cB: out = out & "bin.Currency2.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cB = 922337203685477.5807@: r = cA * cB: out = out & "bin.Currency2.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cB = 922337203685477.5807@: r = cA / cB: out = out & "bin.Currency2.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cB = 922337203685477.5807@: r = cA \ cB: out = out & "bin.Currency2.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cB = 922337203685477.5807@: r = cA Mod cB: out = out & "bin.Currency2.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cB = 922337203685477.5807@: r = cA ^ cB: out = out & "bin.Currency2.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: vB = Empty: r = cA + vB: out = out & "bin.Currency0.add.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: vB = Empty: r = cA - vB: out = out & "bin.Currency0.sub.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: vB = Empty: r = cA * vB: out = out & "bin.Currency0.mul.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: vB = Empty: r = cA / vB: out = out & "bin.Currency0.div.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: vB = Empty: r = cA \ vB: out = out & "bin.Currency0.idiv.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: vB = Empty: r = cA Mod vB: out = out & "bin.Currency0.mod.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: vB = Empty: r = cA ^ vB: out = out & "bin.Currency0.pow.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: vB = 7: r = cA + vB: out = out & "bin.Currency0.add.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: vB = 7: r = cA - vB: out = out & "bin.Currency0.sub.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: vB = 7: r = cA * vB: out = out & "bin.Currency0.mul.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: vB = 7: r = cA / vB: out = out & "bin.Currency0.div.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: vB = 7: r = cA \ vB: out = out & "bin.Currency0.idiv.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: vB = 7: r = cA Mod vB: out = out & "bin.Currency0.mod.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: vB = 7: r = cA ^ vB: out = out & "bin.Currency0.pow.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: vB = 2147483647: r = cA + vB: out = out & "bin.Currency0.add.Variant2=" & Rz(Err.Number, r) & vbLf
  P028 = out
End Function
Function P029() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: cA = -0.5@: vB = 2147483647: r = cA - vB: out = out & "bin.Currency0.sub.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: vB = 2147483647: r = cA * vB: out = out & "bin.Currency0.mul.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: vB = 2147483647: r = cA / vB: out = out & "bin.Currency0.div.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: vB = 2147483647: r = cA \ vB: out = out & "bin.Currency0.idiv.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: vB = 2147483647: r = cA Mod vB: out = out & "bin.Currency0.mod.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: vB = 2147483647: r = cA ^ vB: out = out & "bin.Currency0.pow.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: vB = Empty: r = cA + vB: out = out & "bin.Currency1.add.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: vB = Empty: r = cA - vB: out = out & "bin.Currency1.sub.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: vB = Empty: r = cA * vB: out = out & "bin.Currency1.mul.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: vB = Empty: r = cA / vB: out = out & "bin.Currency1.div.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: vB = Empty: r = cA \ vB: out = out & "bin.Currency1.idiv.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: vB = Empty: r = cA Mod vB: out = out & "bin.Currency1.mod.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: vB = Empty: r = cA ^ vB: out = out & "bin.Currency1.pow.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: vB = 7: r = cA + vB: out = out & "bin.Currency1.add.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: vB = 7: r = cA - vB: out = out & "bin.Currency1.sub.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: vB = 7: r = cA * vB: out = out & "bin.Currency1.mul.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: vB = 7: r = cA / vB: out = out & "bin.Currency1.div.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: vB = 7: r = cA \ vB: out = out & "bin.Currency1.idiv.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: vB = 7: r = cA Mod vB: out = out & "bin.Currency1.mod.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: vB = 7: r = cA ^ vB: out = out & "bin.Currency1.pow.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: vB = 2147483647: r = cA + vB: out = out & "bin.Currency1.add.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: vB = 2147483647: r = cA - vB: out = out & "bin.Currency1.sub.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: vB = 2147483647: r = cA * vB: out = out & "bin.Currency1.mul.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: vB = 2147483647: r = cA / vB: out = out & "bin.Currency1.div.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: vB = 2147483647: r = cA \ vB: out = out & "bin.Currency1.idiv.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: vB = 2147483647: r = cA Mod vB: out = out & "bin.Currency1.mod.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: vB = 2147483647: r = cA ^ vB: out = out & "bin.Currency1.pow.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: vB = Empty: r = cA + vB: out = out & "bin.Currency2.add.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: vB = Empty: r = cA - vB: out = out & "bin.Currency2.sub.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: vB = Empty: r = cA * vB: out = out & "bin.Currency2.mul.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: vB = Empty: r = cA / vB: out = out & "bin.Currency2.div.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: vB = Empty: r = cA \ vB: out = out & "bin.Currency2.idiv.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: vB = Empty: r = cA Mod vB: out = out & "bin.Currency2.mod.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: vB = Empty: r = cA ^ vB: out = out & "bin.Currency2.pow.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: vB = 7: r = cA + vB: out = out & "bin.Currency2.add.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: vB = 7: r = cA - vB: out = out & "bin.Currency2.sub.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: vB = 7: r = cA * vB: out = out & "bin.Currency2.mul.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: vB = 7: r = cA / vB: out = out & "bin.Currency2.div.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: vB = 7: r = cA \ vB: out = out & "bin.Currency2.idiv.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: vB = 7: r = cA Mod vB: out = out & "bin.Currency2.mod.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: vB = 7: r = cA ^ vB: out = out & "bin.Currency2.pow.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: vB = 2147483647: r = cA + vB: out = out & "bin.Currency2.add.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: vB = 2147483647: r = cA - vB: out = out & "bin.Currency2.sub.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: vB = 2147483647: r = cA * vB: out = out & "bin.Currency2.mul.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: vB = 2147483647: r = cA / vB: out = out & "bin.Currency2.div.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: vB = 2147483647: r = cA \ vB: out = out & "bin.Currency2.idiv.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: vB = 2147483647: r = cA Mod vB: out = out & "bin.Currency2.mod.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: vB = 2147483647: r = cA ^ vB: out = out & "bin.Currency2.pow.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 0: r = vA + bB: out = out & "bin.Variant0.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 0: r = vA - bB: out = out & "bin.Variant0.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 0: r = vA * bB: out = out & "bin.Variant0.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 0: r = vA / bB: out = out & "bin.Variant0.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 0: r = vA \ bB: out = out & "bin.Variant0.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 0: r = vA Mod bB: out = out & "bin.Variant0.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 0: r = vA ^ bB: out = out & "bin.Variant0.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 200: r = vA + bB: out = out & "bin.Variant0.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 200: r = vA - bB: out = out & "bin.Variant0.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 200: r = vA * bB: out = out & "bin.Variant0.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 200: r = vA / bB: out = out & "bin.Variant0.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 200: r = vA \ bB: out = out & "bin.Variant0.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 200: r = vA Mod bB: out = out & "bin.Variant0.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 200: r = vA ^ bB: out = out & "bin.Variant0.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 255: r = vA + bB: out = out & "bin.Variant0.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 255: r = vA - bB: out = out & "bin.Variant0.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 255: r = vA * bB: out = out & "bin.Variant0.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 255: r = vA / bB: out = out & "bin.Variant0.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 255: r = vA \ bB: out = out & "bin.Variant0.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 255: r = vA Mod bB: out = out & "bin.Variant0.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 255: r = vA ^ bB: out = out & "bin.Variant0.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: bB = 0: r = vA + bB: out = out & "bin.Variant1.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: bB = 0: r = vA - bB: out = out & "bin.Variant1.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: bB = 0: r = vA * bB: out = out & "bin.Variant1.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: bB = 0: r = vA / bB: out = out & "bin.Variant1.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: bB = 0: r = vA \ bB: out = out & "bin.Variant1.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: bB = 0: r = vA Mod bB: out = out & "bin.Variant1.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: bB = 0: r = vA ^ bB: out = out & "bin.Variant1.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: bB = 200: r = vA + bB: out = out & "bin.Variant1.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: bB = 200: r = vA - bB: out = out & "bin.Variant1.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: bB = 200: r = vA * bB: out = out & "bin.Variant1.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: bB = 200: r = vA / bB: out = out & "bin.Variant1.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: bB = 200: r = vA \ bB: out = out & "bin.Variant1.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: bB = 200: r = vA Mod bB: out = out & "bin.Variant1.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: bB = 200: r = vA ^ bB: out = out & "bin.Variant1.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: bB = 255: r = vA + bB: out = out & "bin.Variant1.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: bB = 255: r = vA - bB: out = out & "bin.Variant1.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: bB = 255: r = vA * bB: out = out & "bin.Variant1.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: bB = 255: r = vA / bB: out = out & "bin.Variant1.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: bB = 255: r = vA \ bB: out = out & "bin.Variant1.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: bB = 255: r = vA Mod bB: out = out & "bin.Variant1.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: bB = 255: r = vA ^ bB: out = out & "bin.Variant1.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 0: r = vA + bB: out = out & "bin.Variant2.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 0: r = vA - bB: out = out & "bin.Variant2.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 0: r = vA * bB: out = out & "bin.Variant2.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 0: r = vA / bB: out = out & "bin.Variant2.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 0: r = vA \ bB: out = out & "bin.Variant2.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 0: r = vA Mod bB: out = out & "bin.Variant2.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 0: r = vA ^ bB: out = out & "bin.Variant2.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 200: r = vA + bB: out = out & "bin.Variant2.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 200: r = vA - bB: out = out & "bin.Variant2.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 200: r = vA * bB: out = out & "bin.Variant2.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 200: r = vA / bB: out = out & "bin.Variant2.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 200: r = vA \ bB: out = out & "bin.Variant2.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 200: r = vA Mod bB: out = out & "bin.Variant2.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 200: r = vA ^ bB: out = out & "bin.Variant2.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 255: r = vA + bB: out = out & "bin.Variant2.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 255: r = vA - bB: out = out & "bin.Variant2.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 255: r = vA * bB: out = out & "bin.Variant2.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 255: r = vA / bB: out = out & "bin.Variant2.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 255: r = vA \ bB: out = out & "bin.Variant2.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 255: r = vA Mod bB: out = out & "bin.Variant2.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 255: r = vA ^ bB: out = out & "bin.Variant2.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = -32768: r = vA + iB: out = out & "bin.Variant0.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = -32768: r = vA - iB: out = out & "bin.Variant0.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = -32768: r = vA * iB: out = out & "bin.Variant0.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = -32768: r = vA / iB: out = out & "bin.Variant0.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = -32768: r = vA \ iB: out = out & "bin.Variant0.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = -32768: r = vA Mod iB: out = out & "bin.Variant0.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = -32768: r = vA ^ iB: out = out & "bin.Variant0.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 1: r = vA + iB: out = out & "bin.Variant0.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 1: r = vA - iB: out = out & "bin.Variant0.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  P029 = out
End Function
Function P030() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: vA = Empty: iB = 1: r = vA * iB: out = out & "bin.Variant0.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 1: r = vA / iB: out = out & "bin.Variant0.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 1: r = vA \ iB: out = out & "bin.Variant0.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 1: r = vA Mod iB: out = out & "bin.Variant0.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 1: r = vA ^ iB: out = out & "bin.Variant0.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 32767: r = vA + iB: out = out & "bin.Variant0.add.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 32767: r = vA - iB: out = out & "bin.Variant0.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 32767: r = vA * iB: out = out & "bin.Variant0.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 32767: r = vA / iB: out = out & "bin.Variant0.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 32767: r = vA \ iB: out = out & "bin.Variant0.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 32767: r = vA Mod iB: out = out & "bin.Variant0.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 32767: r = vA ^ iB: out = out & "bin.Variant0.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: iB = -32768: r = vA + iB: out = out & "bin.Variant1.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: iB = -32768: r = vA - iB: out = out & "bin.Variant1.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: iB = -32768: r = vA * iB: out = out & "bin.Variant1.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: iB = -32768: r = vA / iB: out = out & "bin.Variant1.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: iB = -32768: r = vA \ iB: out = out & "bin.Variant1.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: iB = -32768: r = vA Mod iB: out = out & "bin.Variant1.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: iB = -32768: r = vA ^ iB: out = out & "bin.Variant1.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: iB = 1: r = vA + iB: out = out & "bin.Variant1.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: iB = 1: r = vA - iB: out = out & "bin.Variant1.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: iB = 1: r = vA * iB: out = out & "bin.Variant1.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: iB = 1: r = vA / iB: out = out & "bin.Variant1.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: iB = 1: r = vA \ iB: out = out & "bin.Variant1.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: iB = 1: r = vA Mod iB: out = out & "bin.Variant1.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: iB = 1: r = vA ^ iB: out = out & "bin.Variant1.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: iB = 32767: r = vA + iB: out = out & "bin.Variant1.add.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: iB = 32767: r = vA - iB: out = out & "bin.Variant1.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: iB = 32767: r = vA * iB: out = out & "bin.Variant1.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: iB = 32767: r = vA / iB: out = out & "bin.Variant1.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: iB = 32767: r = vA \ iB: out = out & "bin.Variant1.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: iB = 32767: r = vA Mod iB: out = out & "bin.Variant1.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: iB = 32767: r = vA ^ iB: out = out & "bin.Variant1.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = -32768: r = vA + iB: out = out & "bin.Variant2.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = -32768: r = vA - iB: out = out & "bin.Variant2.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = -32768: r = vA * iB: out = out & "bin.Variant2.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = -32768: r = vA / iB: out = out & "bin.Variant2.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = -32768: r = vA \ iB: out = out & "bin.Variant2.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = -32768: r = vA Mod iB: out = out & "bin.Variant2.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = -32768: r = vA ^ iB: out = out & "bin.Variant2.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 1: r = vA + iB: out = out & "bin.Variant2.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 1: r = vA - iB: out = out & "bin.Variant2.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 1: r = vA * iB: out = out & "bin.Variant2.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 1: r = vA / iB: out = out & "bin.Variant2.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 1: r = vA \ iB: out = out & "bin.Variant2.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 1: r = vA Mod iB: out = out & "bin.Variant2.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 1: r = vA ^ iB: out = out & "bin.Variant2.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 32767: r = vA + iB: out = out & "bin.Variant2.add.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 32767: r = vA - iB: out = out & "bin.Variant2.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 32767: r = vA * iB: out = out & "bin.Variant2.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 32767: r = vA / iB: out = out & "bin.Variant2.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 32767: r = vA \ iB: out = out & "bin.Variant2.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 32767: r = vA Mod iB: out = out & "bin.Variant2.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 32767: r = vA ^ iB: out = out & "bin.Variant2.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = -2147483648: r = vA + lB: out = out & "bin.Variant0.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = -2147483648: r = vA - lB: out = out & "bin.Variant0.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = -2147483648: r = vA * lB: out = out & "bin.Variant0.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = -2147483648: r = vA / lB: out = out & "bin.Variant0.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = -2147483648: r = vA \ lB: out = out & "bin.Variant0.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = -2147483648: r = vA Mod lB: out = out & "bin.Variant0.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = -2147483648: r = vA ^ lB: out = out & "bin.Variant0.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 65536: r = vA + lB: out = out & "bin.Variant0.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 65536: r = vA - lB: out = out & "bin.Variant0.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 65536: r = vA * lB: out = out & "bin.Variant0.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 65536: r = vA / lB: out = out & "bin.Variant0.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 65536: r = vA \ lB: out = out & "bin.Variant0.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 65536: r = vA Mod lB: out = out & "bin.Variant0.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 65536: r = vA ^ lB: out = out & "bin.Variant0.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 2147483647: r = vA + lB: out = out & "bin.Variant0.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 2147483647: r = vA - lB: out = out & "bin.Variant0.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 2147483647: r = vA * lB: out = out & "bin.Variant0.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 2147483647: r = vA / lB: out = out & "bin.Variant0.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 2147483647: r = vA \ lB: out = out & "bin.Variant0.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 2147483647: r = vA Mod lB: out = out & "bin.Variant0.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 2147483647: r = vA ^ lB: out = out & "bin.Variant0.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: lB = -2147483648: r = vA + lB: out = out & "bin.Variant1.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: lB = -2147483648: r = vA - lB: out = out & "bin.Variant1.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: lB = -2147483648: r = vA * lB: out = out & "bin.Variant1.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: lB = -2147483648: r = vA / lB: out = out & "bin.Variant1.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: lB = -2147483648: r = vA \ lB: out = out & "bin.Variant1.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: lB = -2147483648: r = vA Mod lB: out = out & "bin.Variant1.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: lB = -2147483648: r = vA ^ lB: out = out & "bin.Variant1.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: lB = 65536: r = vA + lB: out = out & "bin.Variant1.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: lB = 65536: r = vA - lB: out = out & "bin.Variant1.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: lB = 65536: r = vA * lB: out = out & "bin.Variant1.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: lB = 65536: r = vA / lB: out = out & "bin.Variant1.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: lB = 65536: r = vA \ lB: out = out & "bin.Variant1.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: lB = 65536: r = vA Mod lB: out = out & "bin.Variant1.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: lB = 65536: r = vA ^ lB: out = out & "bin.Variant1.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: lB = 2147483647: r = vA + lB: out = out & "bin.Variant1.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: lB = 2147483647: r = vA - lB: out = out & "bin.Variant1.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: lB = 2147483647: r = vA * lB: out = out & "bin.Variant1.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: lB = 2147483647: r = vA / lB: out = out & "bin.Variant1.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: lB = 2147483647: r = vA \ lB: out = out & "bin.Variant1.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: lB = 2147483647: r = vA Mod lB: out = out & "bin.Variant1.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: lB = 2147483647: r = vA ^ lB: out = out & "bin.Variant1.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = -2147483648: r = vA + lB: out = out & "bin.Variant2.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = -2147483648: r = vA - lB: out = out & "bin.Variant2.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = -2147483648: r = vA * lB: out = out & "bin.Variant2.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = -2147483648: r = vA / lB: out = out & "bin.Variant2.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = -2147483648: r = vA \ lB: out = out & "bin.Variant2.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = -2147483648: r = vA Mod lB: out = out & "bin.Variant2.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = -2147483648: r = vA ^ lB: out = out & "bin.Variant2.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 65536: r = vA + lB: out = out & "bin.Variant2.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 65536: r = vA - lB: out = out & "bin.Variant2.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 65536: r = vA * lB: out = out & "bin.Variant2.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 65536: r = vA / lB: out = out & "bin.Variant2.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 65536: r = vA \ lB: out = out & "bin.Variant2.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 65536: r = vA Mod lB: out = out & "bin.Variant2.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 65536: r = vA ^ lB: out = out & "bin.Variant2.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 2147483647: r = vA + lB: out = out & "bin.Variant2.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 2147483647: r = vA - lB: out = out & "bin.Variant2.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 2147483647: r = vA * lB: out = out & "bin.Variant2.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 2147483647: r = vA / lB: out = out & "bin.Variant2.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 2147483647: r = vA \ lB: out = out & "bin.Variant2.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 2147483647: r = vA Mod lB: out = out & "bin.Variant2.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 2147483647: r = vA ^ lB: out = out & "bin.Variant2.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = -9223372036854775807^: r = vA + qB: out = out & "bin.Variant0.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = -9223372036854775807^: r = vA - qB: out = out & "bin.Variant0.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = -9223372036854775807^: r = vA * qB: out = out & "bin.Variant0.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  P030 = out
End Function
Function P031() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: vA = Empty: qB = -9223372036854775807^: r = vA / qB: out = out & "bin.Variant0.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = -9223372036854775807^: r = vA \ qB: out = out & "bin.Variant0.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = -9223372036854775807^: r = vA Mod qB: out = out & "bin.Variant0.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = -9223372036854775807^: r = vA ^ qB: out = out & "bin.Variant0.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 3: r = vA + qB: out = out & "bin.Variant0.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 3: r = vA - qB: out = out & "bin.Variant0.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 3: r = vA * qB: out = out & "bin.Variant0.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 3: r = vA / qB: out = out & "bin.Variant0.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 3: r = vA \ qB: out = out & "bin.Variant0.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 3: r = vA Mod qB: out = out & "bin.Variant0.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 3: r = vA ^ qB: out = out & "bin.Variant0.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 9223372036854775807^: r = vA + qB: out = out & "bin.Variant0.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 9223372036854775807^: r = vA - qB: out = out & "bin.Variant0.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 9223372036854775807^: r = vA * qB: out = out & "bin.Variant0.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 9223372036854775807^: r = vA / qB: out = out & "bin.Variant0.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 9223372036854775807^: r = vA \ qB: out = out & "bin.Variant0.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 9223372036854775807^: r = vA Mod qB: out = out & "bin.Variant0.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 9223372036854775807^: r = vA ^ qB: out = out & "bin.Variant0.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: qB = -9223372036854775807^: r = vA + qB: out = out & "bin.Variant1.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: qB = -9223372036854775807^: r = vA - qB: out = out & "bin.Variant1.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: qB = -9223372036854775807^: r = vA * qB: out = out & "bin.Variant1.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: qB = -9223372036854775807^: r = vA / qB: out = out & "bin.Variant1.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: qB = -9223372036854775807^: r = vA \ qB: out = out & "bin.Variant1.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: qB = -9223372036854775807^: r = vA Mod qB: out = out & "bin.Variant1.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: qB = -9223372036854775807^: r = vA ^ qB: out = out & "bin.Variant1.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: qB = 3: r = vA + qB: out = out & "bin.Variant1.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: qB = 3: r = vA - qB: out = out & "bin.Variant1.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: qB = 3: r = vA * qB: out = out & "bin.Variant1.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: qB = 3: r = vA / qB: out = out & "bin.Variant1.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: qB = 3: r = vA \ qB: out = out & "bin.Variant1.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: qB = 3: r = vA Mod qB: out = out & "bin.Variant1.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: qB = 3: r = vA ^ qB: out = out & "bin.Variant1.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: qB = 9223372036854775807^: r = vA + qB: out = out & "bin.Variant1.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: qB = 9223372036854775807^: r = vA - qB: out = out & "bin.Variant1.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: qB = 9223372036854775807^: r = vA * qB: out = out & "bin.Variant1.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: qB = 9223372036854775807^: r = vA / qB: out = out & "bin.Variant1.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: qB = 9223372036854775807^: r = vA \ qB: out = out & "bin.Variant1.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: qB = 9223372036854775807^: r = vA Mod qB: out = out & "bin.Variant1.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: qB = 9223372036854775807^: r = vA ^ qB: out = out & "bin.Variant1.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = -9223372036854775807^: r = vA + qB: out = out & "bin.Variant2.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = -9223372036854775807^: r = vA - qB: out = out & "bin.Variant2.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = -9223372036854775807^: r = vA * qB: out = out & "bin.Variant2.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = -9223372036854775807^: r = vA / qB: out = out & "bin.Variant2.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = -9223372036854775807^: r = vA \ qB: out = out & "bin.Variant2.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = -9223372036854775807^: r = vA Mod qB: out = out & "bin.Variant2.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = -9223372036854775807^: r = vA ^ qB: out = out & "bin.Variant2.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 3: r = vA + qB: out = out & "bin.Variant2.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 3: r = vA - qB: out = out & "bin.Variant2.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 3: r = vA * qB: out = out & "bin.Variant2.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 3: r = vA / qB: out = out & "bin.Variant2.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 3: r = vA \ qB: out = out & "bin.Variant2.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 3: r = vA Mod qB: out = out & "bin.Variant2.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 3: r = vA ^ qB: out = out & "bin.Variant2.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 9223372036854775807^: r = vA + qB: out = out & "bin.Variant2.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 9223372036854775807^: r = vA - qB: out = out & "bin.Variant2.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 9223372036854775807^: r = vA * qB: out = out & "bin.Variant2.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 9223372036854775807^: r = vA / qB: out = out & "bin.Variant2.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 9223372036854775807^: r = vA \ qB: out = out & "bin.Variant2.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 9223372036854775807^: r = vA Mod qB: out = out & "bin.Variant2.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 9223372036854775807^: r = vA ^ qB: out = out & "bin.Variant2.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 0.5: r = vA + sB: out = out & "bin.Variant0.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 0.5: r = vA - sB: out = out & "bin.Variant0.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 0.5: r = vA * sB: out = out & "bin.Variant0.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 0.5: r = vA / sB: out = out & "bin.Variant0.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 0.5: r = vA \ sB: out = out & "bin.Variant0.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 0.5: r = vA Mod sB: out = out & "bin.Variant0.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 0.5: r = vA ^ sB: out = out & "bin.Variant0.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 16777216: r = vA + sB: out = out & "bin.Variant0.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 16777216: r = vA - sB: out = out & "bin.Variant0.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 16777216: r = vA * sB: out = out & "bin.Variant0.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 16777216: r = vA / sB: out = out & "bin.Variant0.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 16777216: r = vA \ sB: out = out & "bin.Variant0.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 16777216: r = vA Mod sB: out = out & "bin.Variant0.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 16777216: r = vA ^ sB: out = out & "bin.Variant0.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 3.0E+38: r = vA + sB: out = out & "bin.Variant0.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 3.0E+38: r = vA - sB: out = out & "bin.Variant0.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 3.0E+38: r = vA * sB: out = out & "bin.Variant0.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 3.0E+38: r = vA / sB: out = out & "bin.Variant0.div.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 3.0E+38: r = vA \ sB: out = out & "bin.Variant0.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 3.0E+38: r = vA Mod sB: out = out & "bin.Variant0.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 3.0E+38: r = vA ^ sB: out = out & "bin.Variant0.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: sB = 0.5: r = vA + sB: out = out & "bin.Variant1.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: sB = 0.5: r = vA - sB: out = out & "bin.Variant1.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: sB = 0.5: r = vA * sB: out = out & "bin.Variant1.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: sB = 0.5: r = vA / sB: out = out & "bin.Variant1.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: sB = 0.5: r = vA \ sB: out = out & "bin.Variant1.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: sB = 0.5: r = vA Mod sB: out = out & "bin.Variant1.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: sB = 0.5: r = vA ^ sB: out = out & "bin.Variant1.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: sB = 16777216: r = vA + sB: out = out & "bin.Variant1.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: sB = 16777216: r = vA - sB: out = out & "bin.Variant1.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: sB = 16777216: r = vA * sB: out = out & "bin.Variant1.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: sB = 16777216: r = vA / sB: out = out & "bin.Variant1.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: sB = 16777216: r = vA \ sB: out = out & "bin.Variant1.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: sB = 16777216: r = vA Mod sB: out = out & "bin.Variant1.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: sB = 16777216: r = vA ^ sB: out = out & "bin.Variant1.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: sB = 3.0E+38: r = vA + sB: out = out & "bin.Variant1.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: sB = 3.0E+38: r = vA - sB: out = out & "bin.Variant1.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: sB = 3.0E+38: r = vA * sB: out = out & "bin.Variant1.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: sB = 3.0E+38: r = vA / sB: out = out & "bin.Variant1.div.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: sB = 3.0E+38: r = vA \ sB: out = out & "bin.Variant1.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: sB = 3.0E+38: r = vA Mod sB: out = out & "bin.Variant1.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: sB = 3.0E+38: r = vA ^ sB: out = out & "bin.Variant1.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 0.5: r = vA + sB: out = out & "bin.Variant2.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 0.5: r = vA - sB: out = out & "bin.Variant2.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 0.5: r = vA * sB: out = out & "bin.Variant2.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 0.5: r = vA / sB: out = out & "bin.Variant2.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 0.5: r = vA \ sB: out = out & "bin.Variant2.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 0.5: r = vA Mod sB: out = out & "bin.Variant2.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 0.5: r = vA ^ sB: out = out & "bin.Variant2.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 16777216: r = vA + sB: out = out & "bin.Variant2.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 16777216: r = vA - sB: out = out & "bin.Variant2.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 16777216: r = vA * sB: out = out & "bin.Variant2.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 16777216: r = vA / sB: out = out & "bin.Variant2.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 16777216: r = vA \ sB: out = out & "bin.Variant2.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 16777216: r = vA Mod sB: out = out & "bin.Variant2.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 16777216: r = vA ^ sB: out = out & "bin.Variant2.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 3.0E+38: r = vA + sB: out = out & "bin.Variant2.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 3.0E+38: r = vA - sB: out = out & "bin.Variant2.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 3.0E+38: r = vA * sB: out = out & "bin.Variant2.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 3.0E+38: r = vA / sB: out = out & "bin.Variant2.div.Single2=" & Rz(Err.Number, r) & vbLf
  P031 = out
End Function
Function P032() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: vA = 2147483647: sB = 3.0E+38: r = vA \ sB: out = out & "bin.Variant2.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 3.0E+38: r = vA Mod sB: out = out & "bin.Variant2.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 3.0E+38: r = vA ^ sB: out = out & "bin.Variant2.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 0.5: r = vA + dB: out = out & "bin.Variant0.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 0.5: r = vA - dB: out = out & "bin.Variant0.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 0.5: r = vA * dB: out = out & "bin.Variant0.mul.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 0.5: r = vA / dB: out = out & "bin.Variant0.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 0.5: r = vA \ dB: out = out & "bin.Variant0.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 0.5: r = vA Mod dB: out = out & "bin.Variant0.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 0.5: r = vA ^ dB: out = out & "bin.Variant0.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 2.5: r = vA + dB: out = out & "bin.Variant0.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 2.5: r = vA - dB: out = out & "bin.Variant0.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 2.5: r = vA * dB: out = out & "bin.Variant0.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 2.5: r = vA / dB: out = out & "bin.Variant0.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 2.5: r = vA \ dB: out = out & "bin.Variant0.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 2.5: r = vA Mod dB: out = out & "bin.Variant0.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 2.5: r = vA ^ dB: out = out & "bin.Variant0.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 1.0E+308: r = vA + dB: out = out & "bin.Variant0.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 1.0E+308: r = vA - dB: out = out & "bin.Variant0.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 1.0E+308: r = vA * dB: out = out & "bin.Variant0.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 1.0E+308: r = vA / dB: out = out & "bin.Variant0.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 1.0E+308: r = vA \ dB: out = out & "bin.Variant0.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 1.0E+308: r = vA Mod dB: out = out & "bin.Variant0.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 1.0E+308: r = vA ^ dB: out = out & "bin.Variant0.pow.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: dB = 0.5: r = vA + dB: out = out & "bin.Variant1.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: dB = 0.5: r = vA - dB: out = out & "bin.Variant1.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: dB = 0.5: r = vA * dB: out = out & "bin.Variant1.mul.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: dB = 0.5: r = vA / dB: out = out & "bin.Variant1.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: dB = 0.5: r = vA \ dB: out = out & "bin.Variant1.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: dB = 0.5: r = vA Mod dB: out = out & "bin.Variant1.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: dB = 0.5: r = vA ^ dB: out = out & "bin.Variant1.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: dB = 2.5: r = vA + dB: out = out & "bin.Variant1.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: dB = 2.5: r = vA - dB: out = out & "bin.Variant1.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: dB = 2.5: r = vA * dB: out = out & "bin.Variant1.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: dB = 2.5: r = vA / dB: out = out & "bin.Variant1.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: dB = 2.5: r = vA \ dB: out = out & "bin.Variant1.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: dB = 2.5: r = vA Mod dB: out = out & "bin.Variant1.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: dB = 2.5: r = vA ^ dB: out = out & "bin.Variant1.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: dB = 1.0E+308: r = vA + dB: out = out & "bin.Variant1.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: dB = 1.0E+308: r = vA - dB: out = out & "bin.Variant1.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: dB = 1.0E+308: r = vA * dB: out = out & "bin.Variant1.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: dB = 1.0E+308: r = vA / dB: out = out & "bin.Variant1.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: dB = 1.0E+308: r = vA \ dB: out = out & "bin.Variant1.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: dB = 1.0E+308: r = vA Mod dB: out = out & "bin.Variant1.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: dB = 1.0E+308: r = vA ^ dB: out = out & "bin.Variant1.pow.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 0.5: r = vA + dB: out = out & "bin.Variant2.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 0.5: r = vA - dB: out = out & "bin.Variant2.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 0.5: r = vA * dB: out = out & "bin.Variant2.mul.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 0.5: r = vA / dB: out = out & "bin.Variant2.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 0.5: r = vA \ dB: out = out & "bin.Variant2.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 0.5: r = vA Mod dB: out = out & "bin.Variant2.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 0.5: r = vA ^ dB: out = out & "bin.Variant2.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 2.5: r = vA + dB: out = out & "bin.Variant2.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 2.5: r = vA - dB: out = out & "bin.Variant2.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 2.5: r = vA * dB: out = out & "bin.Variant2.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 2.5: r = vA / dB: out = out & "bin.Variant2.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 2.5: r = vA \ dB: out = out & "bin.Variant2.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 2.5: r = vA Mod dB: out = out & "bin.Variant2.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 2.5: r = vA ^ dB: out = out & "bin.Variant2.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 1.0E+308: r = vA + dB: out = out & "bin.Variant2.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 1.0E+308: r = vA - dB: out = out & "bin.Variant2.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 1.0E+308: r = vA * dB: out = out & "bin.Variant2.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 1.0E+308: r = vA / dB: out = out & "bin.Variant2.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 1.0E+308: r = vA \ dB: out = out & "bin.Variant2.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 1.0E+308: r = vA Mod dB: out = out & "bin.Variant2.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 1.0E+308: r = vA ^ dB: out = out & "bin.Variant2.pow.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = -0.5@: r = vA + cB: out = out & "bin.Variant0.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = -0.5@: r = vA - cB: out = out & "bin.Variant0.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = -0.5@: r = vA * cB: out = out & "bin.Variant0.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = -0.5@: r = vA / cB: out = out & "bin.Variant0.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = -0.5@: r = vA \ cB: out = out & "bin.Variant0.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = -0.5@: r = vA Mod cB: out = out & "bin.Variant0.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = -0.5@: r = vA ^ cB: out = out & "bin.Variant0.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 1.5@: r = vA + cB: out = out & "bin.Variant0.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 1.5@: r = vA - cB: out = out & "bin.Variant0.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 1.5@: r = vA * cB: out = out & "bin.Variant0.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 1.5@: r = vA / cB: out = out & "bin.Variant0.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 1.5@: r = vA \ cB: out = out & "bin.Variant0.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 1.5@: r = vA Mod cB: out = out & "bin.Variant0.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 1.5@: r = vA ^ cB: out = out & "bin.Variant0.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 922337203685477.5807@: r = vA + cB: out = out & "bin.Variant0.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 922337203685477.5807@: r = vA - cB: out = out & "bin.Variant0.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 922337203685477.5807@: r = vA * cB: out = out & "bin.Variant0.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 922337203685477.5807@: r = vA / cB: out = out & "bin.Variant0.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 922337203685477.5807@: r = vA \ cB: out = out & "bin.Variant0.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 922337203685477.5807@: r = vA Mod cB: out = out & "bin.Variant0.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 922337203685477.5807@: r = vA ^ cB: out = out & "bin.Variant0.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: cB = -0.5@: r = vA + cB: out = out & "bin.Variant1.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: cB = -0.5@: r = vA - cB: out = out & "bin.Variant1.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: cB = -0.5@: r = vA * cB: out = out & "bin.Variant1.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: cB = -0.5@: r = vA / cB: out = out & "bin.Variant1.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: cB = -0.5@: r = vA \ cB: out = out & "bin.Variant1.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: cB = -0.5@: r = vA Mod cB: out = out & "bin.Variant1.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: cB = -0.5@: r = vA ^ cB: out = out & "bin.Variant1.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: cB = 1.5@: r = vA + cB: out = out & "bin.Variant1.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: cB = 1.5@: r = vA - cB: out = out & "bin.Variant1.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: cB = 1.5@: r = vA * cB: out = out & "bin.Variant1.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: cB = 1.5@: r = vA / cB: out = out & "bin.Variant1.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: cB = 1.5@: r = vA \ cB: out = out & "bin.Variant1.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: cB = 1.5@: r = vA Mod cB: out = out & "bin.Variant1.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: cB = 1.5@: r = vA ^ cB: out = out & "bin.Variant1.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: cB = 922337203685477.5807@: r = vA + cB: out = out & "bin.Variant1.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: cB = 922337203685477.5807@: r = vA - cB: out = out & "bin.Variant1.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: cB = 922337203685477.5807@: r = vA * cB: out = out & "bin.Variant1.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: cB = 922337203685477.5807@: r = vA / cB: out = out & "bin.Variant1.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: cB = 922337203685477.5807@: r = vA \ cB: out = out & "bin.Variant1.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: cB = 922337203685477.5807@: r = vA Mod cB: out = out & "bin.Variant1.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: cB = 922337203685477.5807@: r = vA ^ cB: out = out & "bin.Variant1.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = -0.5@: r = vA + cB: out = out & "bin.Variant2.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = -0.5@: r = vA - cB: out = out & "bin.Variant2.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = -0.5@: r = vA * cB: out = out & "bin.Variant2.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = -0.5@: r = vA / cB: out = out & "bin.Variant2.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = -0.5@: r = vA \ cB: out = out & "bin.Variant2.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = -0.5@: r = vA Mod cB: out = out & "bin.Variant2.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = -0.5@: r = vA ^ cB: out = out & "bin.Variant2.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 1.5@: r = vA + cB: out = out & "bin.Variant2.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 1.5@: r = vA - cB: out = out & "bin.Variant2.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 1.5@: r = vA * cB: out = out & "bin.Variant2.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 1.5@: r = vA / cB: out = out & "bin.Variant2.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 1.5@: r = vA \ cB: out = out & "bin.Variant2.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  P032 = out
End Function
Function P033() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: vA = 2147483647: cB = 1.5@: r = vA Mod cB: out = out & "bin.Variant2.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 1.5@: r = vA ^ cB: out = out & "bin.Variant2.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 922337203685477.5807@: r = vA + cB: out = out & "bin.Variant2.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 922337203685477.5807@: r = vA - cB: out = out & "bin.Variant2.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 922337203685477.5807@: r = vA * cB: out = out & "bin.Variant2.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 922337203685477.5807@: r = vA / cB: out = out & "bin.Variant2.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 922337203685477.5807@: r = vA \ cB: out = out & "bin.Variant2.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 922337203685477.5807@: r = vA Mod cB: out = out & "bin.Variant2.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 922337203685477.5807@: r = vA ^ cB: out = out & "bin.Variant2.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: vB = 7: r = vA + vB: out = out & "bin.Variant0.add.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: vB = 7: r = vA - vB: out = out & "bin.Variant0.sub.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: vB = 7: r = vA * vB: out = out & "bin.Variant0.mul.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: vB = 7: r = vA / vB: out = out & "bin.Variant0.div.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: vB = 7: r = vA \ vB: out = out & "bin.Variant0.idiv.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: vB = 7: r = vA Mod vB: out = out & "bin.Variant0.mod.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: vB = 7: r = vA ^ vB: out = out & "bin.Variant0.pow.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: vB = 2147483647: r = vA + vB: out = out & "bin.Variant0.add.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: vB = 2147483647: r = vA - vB: out = out & "bin.Variant0.sub.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: vB = 2147483647: r = vA * vB: out = out & "bin.Variant0.mul.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: vB = 2147483647: r = vA / vB: out = out & "bin.Variant0.div.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: vB = 2147483647: r = vA \ vB: out = out & "bin.Variant0.idiv.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: vB = 2147483647: r = vA Mod vB: out = out & "bin.Variant0.mod.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: vB = 2147483647: r = vA ^ vB: out = out & "bin.Variant0.pow.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: vB = Empty: r = vA + vB: out = out & "bin.Variant1.add.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: vB = Empty: r = vA - vB: out = out & "bin.Variant1.sub.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: vB = Empty: r = vA * vB: out = out & "bin.Variant1.mul.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: vB = Empty: r = vA / vB: out = out & "bin.Variant1.div.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: vB = Empty: r = vA \ vB: out = out & "bin.Variant1.idiv.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: vB = Empty: r = vA Mod vB: out = out & "bin.Variant1.mod.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: vB = Empty: r = vA ^ vB: out = out & "bin.Variant1.pow.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: vB = 7: r = vA + vB: out = out & "bin.Variant1.add.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: vB = 7: r = vA - vB: out = out & "bin.Variant1.sub.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: vB = 7: r = vA * vB: out = out & "bin.Variant1.mul.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: vB = 7: r = vA / vB: out = out & "bin.Variant1.div.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: vB = 7: r = vA \ vB: out = out & "bin.Variant1.idiv.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: vB = 7: r = vA Mod vB: out = out & "bin.Variant1.mod.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: vB = 7: r = vA ^ vB: out = out & "bin.Variant1.pow.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: vB = 2147483647: r = vA + vB: out = out & "bin.Variant1.add.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: vB = 2147483647: r = vA - vB: out = out & "bin.Variant1.sub.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: vB = 2147483647: r = vA * vB: out = out & "bin.Variant1.mul.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: vB = 2147483647: r = vA / vB: out = out & "bin.Variant1.div.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: vB = 2147483647: r = vA \ vB: out = out & "bin.Variant1.idiv.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: vB = 2147483647: r = vA Mod vB: out = out & "bin.Variant1.mod.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: vB = 2147483647: r = vA ^ vB: out = out & "bin.Variant1.pow.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: vB = Empty: r = vA + vB: out = out & "bin.Variant2.add.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: vB = Empty: r = vA - vB: out = out & "bin.Variant2.sub.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: vB = Empty: r = vA * vB: out = out & "bin.Variant2.mul.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: vB = Empty: r = vA / vB: out = out & "bin.Variant2.div.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: vB = Empty: r = vA \ vB: out = out & "bin.Variant2.idiv.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: vB = Empty: r = vA Mod vB: out = out & "bin.Variant2.mod.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: vB = Empty: r = vA ^ vB: out = out & "bin.Variant2.pow.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: vB = 7: r = vA + vB: out = out & "bin.Variant2.add.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: vB = 7: r = vA - vB: out = out & "bin.Variant2.sub.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: vB = 7: r = vA * vB: out = out & "bin.Variant2.mul.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: vB = 7: r = vA / vB: out = out & "bin.Variant2.div.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: vB = 7: r = vA \ vB: out = out & "bin.Variant2.idiv.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: vB = 7: r = vA Mod vB: out = out & "bin.Variant2.mod.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: vB = 7: r = vA ^ vB: out = out & "bin.Variant2.pow.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: vB = 2147483647: r = vA + vB: out = out & "bin.Variant2.add.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: vB = 2147483647: r = vA - vB: out = out & "bin.Variant2.sub.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: vB = 2147483647: r = vA * vB: out = out & "bin.Variant2.mul.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: vB = 2147483647: r = vA / vB: out = out & "bin.Variant2.div.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: vB = 2147483647: r = vA \ vB: out = out & "bin.Variant2.idiv.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: vB = 2147483647: r = vA Mod vB: out = out & "bin.Variant2.mod.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: vB = 2147483647: r = vA ^ vB: out = out & "bin.Variant2.pow.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 0: r = vA + bB: out = out & "var.vInt.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 0: r = bB + vA: out = out & "var.Byte0.add.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 0: r = vA - bB: out = out & "var.vInt.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 0: r = bB - vA: out = out & "var.Byte0.sub.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 0: r = vA * bB: out = out & "var.vInt.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 0: r = bB * vA: out = out & "var.Byte0.mul.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 0: r = vA / bB: out = out & "var.vInt.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 0: r = bB / vA: out = out & "var.Byte0.div.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 0: r = vA \ bB: out = out & "var.vInt.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 0: r = bB \ vA: out = out & "var.Byte0.idiv.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 0: r = vA Mod bB: out = out & "var.vInt.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 0: r = bB Mod vA: out = out & "var.Byte0.mod.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 0: r = vA ^ bB: out = out & "var.vInt.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 0: r = bB ^ vA: out = out & "var.Byte0.pow.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 200: r = vA + bB: out = out & "var.vInt.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 200: r = bB + vA: out = out & "var.Byte1.add.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 200: r = vA - bB: out = out & "var.vInt.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 200: r = bB - vA: out = out & "var.Byte1.sub.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 200: r = vA * bB: out = out & "var.vInt.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 200: r = bB * vA: out = out & "var.Byte1.mul.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 200: r = vA / bB: out = out & "var.vInt.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 200: r = bB / vA: out = out & "var.Byte1.div.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 200: r = vA \ bB: out = out & "var.vInt.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 200: r = bB \ vA: out = out & "var.Byte1.idiv.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 200: r = vA Mod bB: out = out & "var.vInt.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 200: r = bB Mod vA: out = out & "var.Byte1.mod.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 200: r = vA ^ bB: out = out & "var.vInt.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 200: r = bB ^ vA: out = out & "var.Byte1.pow.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 255: r = vA + bB: out = out & "var.vInt.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 255: r = bB + vA: out = out & "var.Byte2.add.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 255: r = vA - bB: out = out & "var.vInt.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 255: r = bB - vA: out = out & "var.Byte2.sub.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 255: r = vA * bB: out = out & "var.vInt.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 255: r = bB * vA: out = out & "var.Byte2.mul.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 255: r = vA / bB: out = out & "var.vInt.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 255: r = bB / vA: out = out & "var.Byte2.div.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 255: r = vA \ bB: out = out & "var.vInt.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 255: r = bB \ vA: out = out & "var.Byte2.idiv.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 255: r = vA Mod bB: out = out & "var.vInt.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 255: r = bB Mod vA: out = out & "var.Byte2.mod.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 255: r = vA ^ bB: out = out & "var.vInt.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: bB = 255: r = bB ^ vA: out = out & "var.Byte2.pow.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = -32768: r = vA + iB: out = out & "var.vInt.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = -32768: r = iB + vA: out = out & "var.Integer0.add.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = -32768: r = vA - iB: out = out & "var.vInt.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = -32768: r = iB - vA: out = out & "var.Integer0.sub.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = -32768: r = vA * iB: out = out & "var.vInt.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = -32768: r = iB * vA: out = out & "var.Integer0.mul.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = -32768: r = vA / iB: out = out & "var.vInt.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = -32768: r = iB / vA: out = out & "var.Integer0.div.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = -32768: r = vA \ iB: out = out & "var.vInt.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = -32768: r = iB \ vA: out = out & "var.Integer0.idiv.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = -32768: r = vA Mod iB: out = out & "var.vInt.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = -32768: r = iB Mod vA: out = out & "var.Integer0.mod.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = -32768: r = vA ^ iB: out = out & "var.vInt.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  P033 = out
End Function
Function P034() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: vA = 32767: iB = -32768: r = iB ^ vA: out = out & "var.Integer0.pow.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = 1: r = vA + iB: out = out & "var.vInt.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = 1: r = iB + vA: out = out & "var.Integer1.add.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = 1: r = vA - iB: out = out & "var.vInt.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = 1: r = iB - vA: out = out & "var.Integer1.sub.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = 1: r = vA * iB: out = out & "var.vInt.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = 1: r = iB * vA: out = out & "var.Integer1.mul.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = 1: r = vA / iB: out = out & "var.vInt.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = 1: r = iB / vA: out = out & "var.Integer1.div.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = 1: r = vA \ iB: out = out & "var.vInt.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = 1: r = iB \ vA: out = out & "var.Integer1.idiv.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = 1: r = vA Mod iB: out = out & "var.vInt.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = 1: r = iB Mod vA: out = out & "var.Integer1.mod.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = 1: r = vA ^ iB: out = out & "var.vInt.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = 1: r = iB ^ vA: out = out & "var.Integer1.pow.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = 32767: r = vA + iB: out = out & "var.vInt.add.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = 32767: r = iB + vA: out = out & "var.Integer2.add.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = 32767: r = vA - iB: out = out & "var.vInt.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = 32767: r = iB - vA: out = out & "var.Integer2.sub.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = 32767: r = vA * iB: out = out & "var.vInt.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = 32767: r = iB * vA: out = out & "var.Integer2.mul.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = 32767: r = vA / iB: out = out & "var.vInt.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = 32767: r = iB / vA: out = out & "var.Integer2.div.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = 32767: r = vA \ iB: out = out & "var.vInt.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = 32767: r = iB \ vA: out = out & "var.Integer2.idiv.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = 32767: r = vA Mod iB: out = out & "var.vInt.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = 32767: r = iB Mod vA: out = out & "var.Integer2.mod.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = 32767: r = vA ^ iB: out = out & "var.vInt.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: iB = 32767: r = iB ^ vA: out = out & "var.Integer2.pow.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = -2147483648: r = vA + lB: out = out & "var.vInt.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = -2147483648: r = lB + vA: out = out & "var.Long0.add.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = -2147483648: r = vA - lB: out = out & "var.vInt.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = -2147483648: r = lB - vA: out = out & "var.Long0.sub.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = -2147483648: r = vA * lB: out = out & "var.vInt.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = -2147483648: r = lB * vA: out = out & "var.Long0.mul.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = -2147483648: r = vA / lB: out = out & "var.vInt.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = -2147483648: r = lB / vA: out = out & "var.Long0.div.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = -2147483648: r = vA \ lB: out = out & "var.vInt.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = -2147483648: r = lB \ vA: out = out & "var.Long0.idiv.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = -2147483648: r = vA Mod lB: out = out & "var.vInt.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = -2147483648: r = lB Mod vA: out = out & "var.Long0.mod.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = -2147483648: r = vA ^ lB: out = out & "var.vInt.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = -2147483648: r = lB ^ vA: out = out & "var.Long0.pow.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = 65536: r = vA + lB: out = out & "var.vInt.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = 65536: r = lB + vA: out = out & "var.Long1.add.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = 65536: r = vA - lB: out = out & "var.vInt.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = 65536: r = lB - vA: out = out & "var.Long1.sub.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = 65536: r = vA * lB: out = out & "var.vInt.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = 65536: r = lB * vA: out = out & "var.Long1.mul.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = 65536: r = vA / lB: out = out & "var.vInt.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = 65536: r = lB / vA: out = out & "var.Long1.div.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = 65536: r = vA \ lB: out = out & "var.vInt.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = 65536: r = lB \ vA: out = out & "var.Long1.idiv.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = 65536: r = vA Mod lB: out = out & "var.vInt.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = 65536: r = lB Mod vA: out = out & "var.Long1.mod.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = 65536: r = vA ^ lB: out = out & "var.vInt.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = 65536: r = lB ^ vA: out = out & "var.Long1.pow.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = 2147483647: r = vA + lB: out = out & "var.vInt.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = 2147483647: r = lB + vA: out = out & "var.Long2.add.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = 2147483647: r = vA - lB: out = out & "var.vInt.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = 2147483647: r = lB - vA: out = out & "var.Long2.sub.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = 2147483647: r = vA * lB: out = out & "var.vInt.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = 2147483647: r = lB * vA: out = out & "var.Long2.mul.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = 2147483647: r = vA / lB: out = out & "var.vInt.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = 2147483647: r = lB / vA: out = out & "var.Long2.div.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = 2147483647: r = vA \ lB: out = out & "var.vInt.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = 2147483647: r = lB \ vA: out = out & "var.Long2.idiv.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = 2147483647: r = vA Mod lB: out = out & "var.vInt.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = 2147483647: r = lB Mod vA: out = out & "var.Long2.mod.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = 2147483647: r = vA ^ lB: out = out & "var.vInt.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: lB = 2147483647: r = lB ^ vA: out = out & "var.Long2.pow.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = -9223372036854775807^: r = vA + qB: out = out & "var.vInt.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = -9223372036854775807^: r = qB + vA: out = out & "var.LongLong0.add.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = -9223372036854775807^: r = vA - qB: out = out & "var.vInt.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = -9223372036854775807^: r = qB - vA: out = out & "var.LongLong0.sub.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = -9223372036854775807^: r = vA * qB: out = out & "var.vInt.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = -9223372036854775807^: r = qB * vA: out = out & "var.LongLong0.mul.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = -9223372036854775807^: r = vA / qB: out = out & "var.vInt.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = -9223372036854775807^: r = qB / vA: out = out & "var.LongLong0.div.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = -9223372036854775807^: r = vA \ qB: out = out & "var.vInt.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = -9223372036854775807^: r = qB \ vA: out = out & "var.LongLong0.idiv.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = -9223372036854775807^: r = vA Mod qB: out = out & "var.vInt.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = -9223372036854775807^: r = qB Mod vA: out = out & "var.LongLong0.mod.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = -9223372036854775807^: r = vA ^ qB: out = out & "var.vInt.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = -9223372036854775807^: r = qB ^ vA: out = out & "var.LongLong0.pow.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = 3: r = vA + qB: out = out & "var.vInt.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = 3: r = qB + vA: out = out & "var.LongLong1.add.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = 3: r = vA - qB: out = out & "var.vInt.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = 3: r = qB - vA: out = out & "var.LongLong1.sub.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = 3: r = vA * qB: out = out & "var.vInt.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = 3: r = qB * vA: out = out & "var.LongLong1.mul.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = 3: r = vA / qB: out = out & "var.vInt.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = 3: r = qB / vA: out = out & "var.LongLong1.div.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = 3: r = vA \ qB: out = out & "var.vInt.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = 3: r = qB \ vA: out = out & "var.LongLong1.idiv.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = 3: r = vA Mod qB: out = out & "var.vInt.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = 3: r = qB Mod vA: out = out & "var.LongLong1.mod.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = 3: r = vA ^ qB: out = out & "var.vInt.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = 3: r = qB ^ vA: out = out & "var.LongLong1.pow.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = 9223372036854775807^: r = vA + qB: out = out & "var.vInt.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = 9223372036854775807^: r = qB + vA: out = out & "var.LongLong2.add.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = 9223372036854775807^: r = vA - qB: out = out & "var.vInt.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = 9223372036854775807^: r = qB - vA: out = out & "var.LongLong2.sub.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = 9223372036854775807^: r = vA * qB: out = out & "var.vInt.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = 9223372036854775807^: r = qB * vA: out = out & "var.LongLong2.mul.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = 9223372036854775807^: r = vA / qB: out = out & "var.vInt.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = 9223372036854775807^: r = qB / vA: out = out & "var.LongLong2.div.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = 9223372036854775807^: r = vA \ qB: out = out & "var.vInt.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = 9223372036854775807^: r = qB \ vA: out = out & "var.LongLong2.idiv.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = 9223372036854775807^: r = vA Mod qB: out = out & "var.vInt.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = 9223372036854775807^: r = qB Mod vA: out = out & "var.LongLong2.mod.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = 9223372036854775807^: r = vA ^ qB: out = out & "var.vInt.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: qB = 9223372036854775807^: r = qB ^ vA: out = out & "var.LongLong2.pow.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 0.5: r = vA + sB: out = out & "var.vInt.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 0.5: r = sB + vA: out = out & "var.Single0.add.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 0.5: r = vA - sB: out = out & "var.vInt.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 0.5: r = sB - vA: out = out & "var.Single0.sub.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 0.5: r = vA * sB: out = out & "var.vInt.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 0.5: r = sB * vA: out = out & "var.Single0.mul.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 0.5: r = vA / sB: out = out & "var.vInt.div.Single0=" & Rz(Err.Number, r) & vbLf
  P034 = out
End Function
Function P035() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: vA = 32767: sB = 0.5: r = sB / vA: out = out & "var.Single0.div.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 0.5: r = vA \ sB: out = out & "var.vInt.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 0.5: r = sB \ vA: out = out & "var.Single0.idiv.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 0.5: r = vA Mod sB: out = out & "var.vInt.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 0.5: r = sB Mod vA: out = out & "var.Single0.mod.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 0.5: r = vA ^ sB: out = out & "var.vInt.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 0.5: r = sB ^ vA: out = out & "var.Single0.pow.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 16777216: r = vA + sB: out = out & "var.vInt.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 16777216: r = sB + vA: out = out & "var.Single1.add.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 16777216: r = vA - sB: out = out & "var.vInt.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 16777216: r = sB - vA: out = out & "var.Single1.sub.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 16777216: r = vA * sB: out = out & "var.vInt.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 16777216: r = sB * vA: out = out & "var.Single1.mul.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 16777216: r = vA / sB: out = out & "var.vInt.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 16777216: r = sB / vA: out = out & "var.Single1.div.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 16777216: r = vA \ sB: out = out & "var.vInt.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 16777216: r = sB \ vA: out = out & "var.Single1.idiv.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 16777216: r = vA Mod sB: out = out & "var.vInt.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 16777216: r = sB Mod vA: out = out & "var.Single1.mod.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 16777216: r = vA ^ sB: out = out & "var.vInt.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 16777216: r = sB ^ vA: out = out & "var.Single1.pow.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 3.0E+38: r = vA + sB: out = out & "var.vInt.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 3.0E+38: r = sB + vA: out = out & "var.Single2.add.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 3.0E+38: r = vA - sB: out = out & "var.vInt.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 3.0E+38: r = sB - vA: out = out & "var.Single2.sub.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 3.0E+38: r = vA * sB: out = out & "var.vInt.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 3.0E+38: r = sB * vA: out = out & "var.Single2.mul.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 3.0E+38: r = vA / sB: out = out & "var.vInt.div.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 3.0E+38: r = sB / vA: out = out & "var.Single2.div.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 3.0E+38: r = vA \ sB: out = out & "var.vInt.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 3.0E+38: r = sB \ vA: out = out & "var.Single2.idiv.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 3.0E+38: r = vA Mod sB: out = out & "var.vInt.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 3.0E+38: r = sB Mod vA: out = out & "var.Single2.mod.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 3.0E+38: r = vA ^ sB: out = out & "var.vInt.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: sB = 3.0E+38: r = sB ^ vA: out = out & "var.Single2.pow.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 0.5: r = vA + dB: out = out & "var.vInt.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 0.5: r = dB + vA: out = out & "var.Double0.add.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 0.5: r = vA - dB: out = out & "var.vInt.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 0.5: r = dB - vA: out = out & "var.Double0.sub.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 0.5: r = vA * dB: out = out & "var.vInt.mul.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 0.5: r = dB * vA: out = out & "var.Double0.mul.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 0.5: r = vA / dB: out = out & "var.vInt.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 0.5: r = dB / vA: out = out & "var.Double0.div.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 0.5: r = vA \ dB: out = out & "var.vInt.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 0.5: r = dB \ vA: out = out & "var.Double0.idiv.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 0.5: r = vA Mod dB: out = out & "var.vInt.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 0.5: r = dB Mod vA: out = out & "var.Double0.mod.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 0.5: r = vA ^ dB: out = out & "var.vInt.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 0.5: r = dB ^ vA: out = out & "var.Double0.pow.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 2.5: r = vA + dB: out = out & "var.vInt.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 2.5: r = dB + vA: out = out & "var.Double1.add.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 2.5: r = vA - dB: out = out & "var.vInt.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 2.5: r = dB - vA: out = out & "var.Double1.sub.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 2.5: r = vA * dB: out = out & "var.vInt.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 2.5: r = dB * vA: out = out & "var.Double1.mul.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 2.5: r = vA / dB: out = out & "var.vInt.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 2.5: r = dB / vA: out = out & "var.Double1.div.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 2.5: r = vA \ dB: out = out & "var.vInt.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 2.5: r = dB \ vA: out = out & "var.Double1.idiv.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 2.5: r = vA Mod dB: out = out & "var.vInt.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 2.5: r = dB Mod vA: out = out & "var.Double1.mod.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 2.5: r = vA ^ dB: out = out & "var.vInt.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 2.5: r = dB ^ vA: out = out & "var.Double1.pow.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 1.0E+308: r = vA + dB: out = out & "var.vInt.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 1.0E+308: r = dB + vA: out = out & "var.Double2.add.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 1.0E+308: r = vA - dB: out = out & "var.vInt.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 1.0E+308: r = dB - vA: out = out & "var.Double2.sub.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 1.0E+308: r = vA * dB: out = out & "var.vInt.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 1.0E+308: r = dB * vA: out = out & "var.Double2.mul.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 1.0E+308: r = vA / dB: out = out & "var.vInt.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 1.0E+308: r = dB / vA: out = out & "var.Double2.div.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 1.0E+308: r = vA \ dB: out = out & "var.vInt.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 1.0E+308: r = dB \ vA: out = out & "var.Double2.idiv.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 1.0E+308: r = vA Mod dB: out = out & "var.vInt.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 1.0E+308: r = dB Mod vA: out = out & "var.Double2.mod.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 1.0E+308: r = vA ^ dB: out = out & "var.vInt.pow.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: dB = 1.0E+308: r = dB ^ vA: out = out & "var.Double2.pow.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = -0.5@: r = vA + cB: out = out & "var.vInt.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = -0.5@: r = cB + vA: out = out & "var.Currency0.add.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = -0.5@: r = vA - cB: out = out & "var.vInt.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = -0.5@: r = cB - vA: out = out & "var.Currency0.sub.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = -0.5@: r = vA * cB: out = out & "var.vInt.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = -0.5@: r = cB * vA: out = out & "var.Currency0.mul.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = -0.5@: r = vA / cB: out = out & "var.vInt.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = -0.5@: r = cB / vA: out = out & "var.Currency0.div.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = -0.5@: r = vA \ cB: out = out & "var.vInt.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = -0.5@: r = cB \ vA: out = out & "var.Currency0.idiv.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = -0.5@: r = vA Mod cB: out = out & "var.vInt.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = -0.5@: r = cB Mod vA: out = out & "var.Currency0.mod.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = -0.5@: r = vA ^ cB: out = out & "var.vInt.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = -0.5@: r = cB ^ vA: out = out & "var.Currency0.pow.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = 1.5@: r = vA + cB: out = out & "var.vInt.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = 1.5@: r = cB + vA: out = out & "var.Currency1.add.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = 1.5@: r = vA - cB: out = out & "var.vInt.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = 1.5@: r = cB - vA: out = out & "var.Currency1.sub.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = 1.5@: r = vA * cB: out = out & "var.vInt.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = 1.5@: r = cB * vA: out = out & "var.Currency1.mul.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = 1.5@: r = vA / cB: out = out & "var.vInt.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = 1.5@: r = cB / vA: out = out & "var.Currency1.div.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = 1.5@: r = vA \ cB: out = out & "var.vInt.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = 1.5@: r = cB \ vA: out = out & "var.Currency1.idiv.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = 1.5@: r = vA Mod cB: out = out & "var.vInt.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = 1.5@: r = cB Mod vA: out = out & "var.Currency1.mod.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = 1.5@: r = vA ^ cB: out = out & "var.vInt.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = 1.5@: r = cB ^ vA: out = out & "var.Currency1.pow.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = 922337203685477.5807@: r = vA + cB: out = out & "var.vInt.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = 922337203685477.5807@: r = cB + vA: out = out & "var.Currency2.add.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = 922337203685477.5807@: r = vA - cB: out = out & "var.vInt.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = 922337203685477.5807@: r = cB - vA: out = out & "var.Currency2.sub.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = 922337203685477.5807@: r = vA * cB: out = out & "var.vInt.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = 922337203685477.5807@: r = cB * vA: out = out & "var.Currency2.mul.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = 922337203685477.5807@: r = vA / cB: out = out & "var.vInt.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = 922337203685477.5807@: r = cB / vA: out = out & "var.Currency2.div.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = 922337203685477.5807@: r = vA \ cB: out = out & "var.vInt.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = 922337203685477.5807@: r = cB \ vA: out = out & "var.Currency2.idiv.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = 922337203685477.5807@: r = vA Mod cB: out = out & "var.vInt.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = 922337203685477.5807@: r = cB Mod vA: out = out & "var.Currency2.mod.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = 922337203685477.5807@: r = vA ^ cB: out = out & "var.vInt.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: cB = 922337203685477.5807@: r = cB ^ vA: out = out & "var.Currency2.pow.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 0: r = vA + bB: out = out & "var.vLng.add.Byte0=" & Rz(Err.Number, r) & vbLf
  P035 = out
End Function
Function P036() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: vA = 2147483647: bB = 0: r = bB + vA: out = out & "var.Byte0.add.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 0: r = vA - bB: out = out & "var.vLng.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 0: r = bB - vA: out = out & "var.Byte0.sub.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 0: r = vA * bB: out = out & "var.vLng.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 0: r = bB * vA: out = out & "var.Byte0.mul.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 0: r = vA / bB: out = out & "var.vLng.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 0: r = bB / vA: out = out & "var.Byte0.div.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 0: r = vA \ bB: out = out & "var.vLng.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 0: r = bB \ vA: out = out & "var.Byte0.idiv.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 0: r = vA Mod bB: out = out & "var.vLng.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 0: r = bB Mod vA: out = out & "var.Byte0.mod.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 0: r = vA ^ bB: out = out & "var.vLng.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 0: r = bB ^ vA: out = out & "var.Byte0.pow.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 200: r = vA + bB: out = out & "var.vLng.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 200: r = bB + vA: out = out & "var.Byte1.add.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 200: r = vA - bB: out = out & "var.vLng.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 200: r = bB - vA: out = out & "var.Byte1.sub.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 200: r = vA * bB: out = out & "var.vLng.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 200: r = bB * vA: out = out & "var.Byte1.mul.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 200: r = vA / bB: out = out & "var.vLng.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 200: r = bB / vA: out = out & "var.Byte1.div.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 200: r = vA \ bB: out = out & "var.vLng.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 200: r = bB \ vA: out = out & "var.Byte1.idiv.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 200: r = vA Mod bB: out = out & "var.vLng.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 200: r = bB Mod vA: out = out & "var.Byte1.mod.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 200: r = vA ^ bB: out = out & "var.vLng.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 200: r = bB ^ vA: out = out & "var.Byte1.pow.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 255: r = vA + bB: out = out & "var.vLng.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 255: r = bB + vA: out = out & "var.Byte2.add.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 255: r = vA - bB: out = out & "var.vLng.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 255: r = bB - vA: out = out & "var.Byte2.sub.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 255: r = vA * bB: out = out & "var.vLng.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 255: r = bB * vA: out = out & "var.Byte2.mul.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 255: r = vA / bB: out = out & "var.vLng.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 255: r = bB / vA: out = out & "var.Byte2.div.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 255: r = vA \ bB: out = out & "var.vLng.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 255: r = bB \ vA: out = out & "var.Byte2.idiv.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 255: r = vA Mod bB: out = out & "var.vLng.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 255: r = bB Mod vA: out = out & "var.Byte2.mod.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 255: r = vA ^ bB: out = out & "var.vLng.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: bB = 255: r = bB ^ vA: out = out & "var.Byte2.pow.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = -32768: r = vA + iB: out = out & "var.vLng.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = -32768: r = iB + vA: out = out & "var.Integer0.add.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = -32768: r = vA - iB: out = out & "var.vLng.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = -32768: r = iB - vA: out = out & "var.Integer0.sub.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = -32768: r = vA * iB: out = out & "var.vLng.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = -32768: r = iB * vA: out = out & "var.Integer0.mul.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = -32768: r = vA / iB: out = out & "var.vLng.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = -32768: r = iB / vA: out = out & "var.Integer0.div.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = -32768: r = vA \ iB: out = out & "var.vLng.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = -32768: r = iB \ vA: out = out & "var.Integer0.idiv.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = -32768: r = vA Mod iB: out = out & "var.vLng.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = -32768: r = iB Mod vA: out = out & "var.Integer0.mod.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = -32768: r = vA ^ iB: out = out & "var.vLng.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = -32768: r = iB ^ vA: out = out & "var.Integer0.pow.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 1: r = vA + iB: out = out & "var.vLng.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 1: r = iB + vA: out = out & "var.Integer1.add.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 1: r = vA - iB: out = out & "var.vLng.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 1: r = iB - vA: out = out & "var.Integer1.sub.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 1: r = vA * iB: out = out & "var.vLng.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 1: r = iB * vA: out = out & "var.Integer1.mul.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 1: r = vA / iB: out = out & "var.vLng.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 1: r = iB / vA: out = out & "var.Integer1.div.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 1: r = vA \ iB: out = out & "var.vLng.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 1: r = iB \ vA: out = out & "var.Integer1.idiv.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 1: r = vA Mod iB: out = out & "var.vLng.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 1: r = iB Mod vA: out = out & "var.Integer1.mod.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 1: r = vA ^ iB: out = out & "var.vLng.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 1: r = iB ^ vA: out = out & "var.Integer1.pow.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 32767: r = vA + iB: out = out & "var.vLng.add.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 32767: r = iB + vA: out = out & "var.Integer2.add.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 32767: r = vA - iB: out = out & "var.vLng.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 32767: r = iB - vA: out = out & "var.Integer2.sub.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 32767: r = vA * iB: out = out & "var.vLng.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 32767: r = iB * vA: out = out & "var.Integer2.mul.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 32767: r = vA / iB: out = out & "var.vLng.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 32767: r = iB / vA: out = out & "var.Integer2.div.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 32767: r = vA \ iB: out = out & "var.vLng.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 32767: r = iB \ vA: out = out & "var.Integer2.idiv.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 32767: r = vA Mod iB: out = out & "var.vLng.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 32767: r = iB Mod vA: out = out & "var.Integer2.mod.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 32767: r = vA ^ iB: out = out & "var.vLng.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: iB = 32767: r = iB ^ vA: out = out & "var.Integer2.pow.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = -2147483648: r = vA + lB: out = out & "var.vLng.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = -2147483648: r = lB + vA: out = out & "var.Long0.add.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = -2147483648: r = vA - lB: out = out & "var.vLng.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = -2147483648: r = lB - vA: out = out & "var.Long0.sub.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = -2147483648: r = vA * lB: out = out & "var.vLng.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = -2147483648: r = lB * vA: out = out & "var.Long0.mul.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = -2147483648: r = vA / lB: out = out & "var.vLng.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = -2147483648: r = lB / vA: out = out & "var.Long0.div.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = -2147483648: r = vA \ lB: out = out & "var.vLng.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = -2147483648: r = lB \ vA: out = out & "var.Long0.idiv.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = -2147483648: r = vA Mod lB: out = out & "var.vLng.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = -2147483648: r = lB Mod vA: out = out & "var.Long0.mod.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = -2147483648: r = vA ^ lB: out = out & "var.vLng.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = -2147483648: r = lB ^ vA: out = out & "var.Long0.pow.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 65536: r = vA + lB: out = out & "var.vLng.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 65536: r = lB + vA: out = out & "var.Long1.add.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 65536: r = vA - lB: out = out & "var.vLng.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 65536: r = lB - vA: out = out & "var.Long1.sub.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 65536: r = vA * lB: out = out & "var.vLng.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 65536: r = lB * vA: out = out & "var.Long1.mul.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 65536: r = vA / lB: out = out & "var.vLng.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 65536: r = lB / vA: out = out & "var.Long1.div.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 65536: r = vA \ lB: out = out & "var.vLng.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 65536: r = lB \ vA: out = out & "var.Long1.idiv.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 65536: r = vA Mod lB: out = out & "var.vLng.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 65536: r = lB Mod vA: out = out & "var.Long1.mod.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 65536: r = vA ^ lB: out = out & "var.vLng.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 65536: r = lB ^ vA: out = out & "var.Long1.pow.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 2147483647: r = vA + lB: out = out & "var.vLng.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 2147483647: r = lB + vA: out = out & "var.Long2.add.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 2147483647: r = vA - lB: out = out & "var.vLng.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 2147483647: r = lB - vA: out = out & "var.Long2.sub.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 2147483647: r = vA * lB: out = out & "var.vLng.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 2147483647: r = lB * vA: out = out & "var.Long2.mul.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 2147483647: r = vA / lB: out = out & "var.vLng.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 2147483647: r = lB / vA: out = out & "var.Long2.div.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 2147483647: r = vA \ lB: out = out & "var.vLng.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  P036 = out
End Function
Function P037() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: vA = 2147483647: lB = 2147483647: r = lB \ vA: out = out & "var.Long2.idiv.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 2147483647: r = vA Mod lB: out = out & "var.vLng.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 2147483647: r = lB Mod vA: out = out & "var.Long2.mod.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 2147483647: r = vA ^ lB: out = out & "var.vLng.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: lB = 2147483647: r = lB ^ vA: out = out & "var.Long2.pow.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = -9223372036854775807^: r = vA + qB: out = out & "var.vLng.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = -9223372036854775807^: r = qB + vA: out = out & "var.LongLong0.add.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = -9223372036854775807^: r = vA - qB: out = out & "var.vLng.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = -9223372036854775807^: r = qB - vA: out = out & "var.LongLong0.sub.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = -9223372036854775807^: r = vA * qB: out = out & "var.vLng.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = -9223372036854775807^: r = qB * vA: out = out & "var.LongLong0.mul.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = -9223372036854775807^: r = vA / qB: out = out & "var.vLng.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = -9223372036854775807^: r = qB / vA: out = out & "var.LongLong0.div.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = -9223372036854775807^: r = vA \ qB: out = out & "var.vLng.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = -9223372036854775807^: r = qB \ vA: out = out & "var.LongLong0.idiv.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = -9223372036854775807^: r = vA Mod qB: out = out & "var.vLng.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = -9223372036854775807^: r = qB Mod vA: out = out & "var.LongLong0.mod.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = -9223372036854775807^: r = vA ^ qB: out = out & "var.vLng.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = -9223372036854775807^: r = qB ^ vA: out = out & "var.LongLong0.pow.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 3: r = vA + qB: out = out & "var.vLng.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 3: r = qB + vA: out = out & "var.LongLong1.add.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 3: r = vA - qB: out = out & "var.vLng.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 3: r = qB - vA: out = out & "var.LongLong1.sub.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 3: r = vA * qB: out = out & "var.vLng.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 3: r = qB * vA: out = out & "var.LongLong1.mul.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 3: r = vA / qB: out = out & "var.vLng.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 3: r = qB / vA: out = out & "var.LongLong1.div.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 3: r = vA \ qB: out = out & "var.vLng.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 3: r = qB \ vA: out = out & "var.LongLong1.idiv.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 3: r = vA Mod qB: out = out & "var.vLng.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 3: r = qB Mod vA: out = out & "var.LongLong1.mod.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 3: r = vA ^ qB: out = out & "var.vLng.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 3: r = qB ^ vA: out = out & "var.LongLong1.pow.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 9223372036854775807^: r = vA + qB: out = out & "var.vLng.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 9223372036854775807^: r = qB + vA: out = out & "var.LongLong2.add.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 9223372036854775807^: r = vA - qB: out = out & "var.vLng.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 9223372036854775807^: r = qB - vA: out = out & "var.LongLong2.sub.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 9223372036854775807^: r = vA * qB: out = out & "var.vLng.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 9223372036854775807^: r = qB * vA: out = out & "var.LongLong2.mul.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 9223372036854775807^: r = vA / qB: out = out & "var.vLng.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 9223372036854775807^: r = qB / vA: out = out & "var.LongLong2.div.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 9223372036854775807^: r = vA \ qB: out = out & "var.vLng.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 9223372036854775807^: r = qB \ vA: out = out & "var.LongLong2.idiv.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 9223372036854775807^: r = vA Mod qB: out = out & "var.vLng.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 9223372036854775807^: r = qB Mod vA: out = out & "var.LongLong2.mod.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 9223372036854775807^: r = vA ^ qB: out = out & "var.vLng.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: qB = 9223372036854775807^: r = qB ^ vA: out = out & "var.LongLong2.pow.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 0.5: r = vA + sB: out = out & "var.vLng.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 0.5: r = sB + vA: out = out & "var.Single0.add.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 0.5: r = vA - sB: out = out & "var.vLng.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 0.5: r = sB - vA: out = out & "var.Single0.sub.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 0.5: r = vA * sB: out = out & "var.vLng.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 0.5: r = sB * vA: out = out & "var.Single0.mul.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 0.5: r = vA / sB: out = out & "var.vLng.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 0.5: r = sB / vA: out = out & "var.Single0.div.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 0.5: r = vA \ sB: out = out & "var.vLng.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 0.5: r = sB \ vA: out = out & "var.Single0.idiv.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 0.5: r = vA Mod sB: out = out & "var.vLng.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 0.5: r = sB Mod vA: out = out & "var.Single0.mod.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 0.5: r = vA ^ sB: out = out & "var.vLng.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 0.5: r = sB ^ vA: out = out & "var.Single0.pow.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 16777216: r = vA + sB: out = out & "var.vLng.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 16777216: r = sB + vA: out = out & "var.Single1.add.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 16777216: r = vA - sB: out = out & "var.vLng.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 16777216: r = sB - vA: out = out & "var.Single1.sub.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 16777216: r = vA * sB: out = out & "var.vLng.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 16777216: r = sB * vA: out = out & "var.Single1.mul.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 16777216: r = vA / sB: out = out & "var.vLng.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 16777216: r = sB / vA: out = out & "var.Single1.div.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 16777216: r = vA \ sB: out = out & "var.vLng.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 16777216: r = sB \ vA: out = out & "var.Single1.idiv.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 16777216: r = vA Mod sB: out = out & "var.vLng.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 16777216: r = sB Mod vA: out = out & "var.Single1.mod.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 16777216: r = vA ^ sB: out = out & "var.vLng.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 16777216: r = sB ^ vA: out = out & "var.Single1.pow.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 3.0E+38: r = vA + sB: out = out & "var.vLng.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 3.0E+38: r = sB + vA: out = out & "var.Single2.add.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 3.0E+38: r = vA - sB: out = out & "var.vLng.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 3.0E+38: r = sB - vA: out = out & "var.Single2.sub.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 3.0E+38: r = vA * sB: out = out & "var.vLng.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 3.0E+38: r = sB * vA: out = out & "var.Single2.mul.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 3.0E+38: r = vA / sB: out = out & "var.vLng.div.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 3.0E+38: r = sB / vA: out = out & "var.Single2.div.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 3.0E+38: r = vA \ sB: out = out & "var.vLng.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 3.0E+38: r = sB \ vA: out = out & "var.Single2.idiv.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 3.0E+38: r = vA Mod sB: out = out & "var.vLng.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 3.0E+38: r = sB Mod vA: out = out & "var.Single2.mod.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 3.0E+38: r = vA ^ sB: out = out & "var.vLng.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: sB = 3.0E+38: r = sB ^ vA: out = out & "var.Single2.pow.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 0.5: r = vA + dB: out = out & "var.vLng.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 0.5: r = dB + vA: out = out & "var.Double0.add.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 0.5: r = vA - dB: out = out & "var.vLng.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 0.5: r = dB - vA: out = out & "var.Double0.sub.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 0.5: r = vA * dB: out = out & "var.vLng.mul.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 0.5: r = dB * vA: out = out & "var.Double0.mul.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 0.5: r = vA / dB: out = out & "var.vLng.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 0.5: r = dB / vA: out = out & "var.Double0.div.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 0.5: r = vA \ dB: out = out & "var.vLng.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 0.5: r = dB \ vA: out = out & "var.Double0.idiv.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 0.5: r = vA Mod dB: out = out & "var.vLng.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 0.5: r = dB Mod vA: out = out & "var.Double0.mod.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 0.5: r = vA ^ dB: out = out & "var.vLng.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 0.5: r = dB ^ vA: out = out & "var.Double0.pow.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 2.5: r = vA + dB: out = out & "var.vLng.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 2.5: r = dB + vA: out = out & "var.Double1.add.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 2.5: r = vA - dB: out = out & "var.vLng.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 2.5: r = dB - vA: out = out & "var.Double1.sub.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 2.5: r = vA * dB: out = out & "var.vLng.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 2.5: r = dB * vA: out = out & "var.Double1.mul.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 2.5: r = vA / dB: out = out & "var.vLng.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 2.5: r = dB / vA: out = out & "var.Double1.div.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 2.5: r = vA \ dB: out = out & "var.vLng.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 2.5: r = dB \ vA: out = out & "var.Double1.idiv.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 2.5: r = vA Mod dB: out = out & "var.vLng.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 2.5: r = dB Mod vA: out = out & "var.Double1.mod.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 2.5: r = vA ^ dB: out = out & "var.vLng.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 2.5: r = dB ^ vA: out = out & "var.Double1.pow.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 1.0E+308: r = vA + dB: out = out & "var.vLng.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 1.0E+308: r = dB + vA: out = out & "var.Double2.add.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 1.0E+308: r = vA - dB: out = out & "var.vLng.sub.Double2=" & Rz(Err.Number, r) & vbLf
  P037 = out
End Function
Function P038() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: vA = 2147483647: dB = 1.0E+308: r = dB - vA: out = out & "var.Double2.sub.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 1.0E+308: r = vA * dB: out = out & "var.vLng.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 1.0E+308: r = dB * vA: out = out & "var.Double2.mul.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 1.0E+308: r = vA / dB: out = out & "var.vLng.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 1.0E+308: r = dB / vA: out = out & "var.Double2.div.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 1.0E+308: r = vA \ dB: out = out & "var.vLng.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 1.0E+308: r = dB \ vA: out = out & "var.Double2.idiv.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 1.0E+308: r = vA Mod dB: out = out & "var.vLng.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 1.0E+308: r = dB Mod vA: out = out & "var.Double2.mod.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 1.0E+308: r = vA ^ dB: out = out & "var.vLng.pow.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: dB = 1.0E+308: r = dB ^ vA: out = out & "var.Double2.pow.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = -0.5@: r = vA + cB: out = out & "var.vLng.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = -0.5@: r = cB + vA: out = out & "var.Currency0.add.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = -0.5@: r = vA - cB: out = out & "var.vLng.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = -0.5@: r = cB - vA: out = out & "var.Currency0.sub.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = -0.5@: r = vA * cB: out = out & "var.vLng.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = -0.5@: r = cB * vA: out = out & "var.Currency0.mul.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = -0.5@: r = vA / cB: out = out & "var.vLng.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = -0.5@: r = cB / vA: out = out & "var.Currency0.div.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = -0.5@: r = vA \ cB: out = out & "var.vLng.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = -0.5@: r = cB \ vA: out = out & "var.Currency0.idiv.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = -0.5@: r = vA Mod cB: out = out & "var.vLng.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = -0.5@: r = cB Mod vA: out = out & "var.Currency0.mod.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = -0.5@: r = vA ^ cB: out = out & "var.vLng.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = -0.5@: r = cB ^ vA: out = out & "var.Currency0.pow.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 1.5@: r = vA + cB: out = out & "var.vLng.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 1.5@: r = cB + vA: out = out & "var.Currency1.add.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 1.5@: r = vA - cB: out = out & "var.vLng.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 1.5@: r = cB - vA: out = out & "var.Currency1.sub.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 1.5@: r = vA * cB: out = out & "var.vLng.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 1.5@: r = cB * vA: out = out & "var.Currency1.mul.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 1.5@: r = vA / cB: out = out & "var.vLng.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 1.5@: r = cB / vA: out = out & "var.Currency1.div.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 1.5@: r = vA \ cB: out = out & "var.vLng.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 1.5@: r = cB \ vA: out = out & "var.Currency1.idiv.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 1.5@: r = vA Mod cB: out = out & "var.vLng.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 1.5@: r = cB Mod vA: out = out & "var.Currency1.mod.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 1.5@: r = vA ^ cB: out = out & "var.vLng.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 1.5@: r = cB ^ vA: out = out & "var.Currency1.pow.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 922337203685477.5807@: r = vA + cB: out = out & "var.vLng.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 922337203685477.5807@: r = cB + vA: out = out & "var.Currency2.add.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 922337203685477.5807@: r = vA - cB: out = out & "var.vLng.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 922337203685477.5807@: r = cB - vA: out = out & "var.Currency2.sub.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 922337203685477.5807@: r = vA * cB: out = out & "var.vLng.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 922337203685477.5807@: r = cB * vA: out = out & "var.Currency2.mul.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 922337203685477.5807@: r = vA / cB: out = out & "var.vLng.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 922337203685477.5807@: r = cB / vA: out = out & "var.Currency2.div.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 922337203685477.5807@: r = vA \ cB: out = out & "var.vLng.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 922337203685477.5807@: r = cB \ vA: out = out & "var.Currency2.idiv.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 922337203685477.5807@: r = vA Mod cB: out = out & "var.vLng.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 922337203685477.5807@: r = cB Mod vA: out = out & "var.Currency2.mod.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 922337203685477.5807@: r = vA ^ cB: out = out & "var.vLng.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: cB = 922337203685477.5807@: r = cB ^ vA: out = out & "var.Currency2.pow.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 0: r = vA + bB: out = out & "var.vDbl.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 0: r = bB + vA: out = out & "var.Byte0.add.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 0: r = vA - bB: out = out & "var.vDbl.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 0: r = bB - vA: out = out & "var.Byte0.sub.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 0: r = vA * bB: out = out & "var.vDbl.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 0: r = bB * vA: out = out & "var.Byte0.mul.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 0: r = vA / bB: out = out & "var.vDbl.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 0: r = bB / vA: out = out & "var.Byte0.div.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 0: r = vA \ bB: out = out & "var.vDbl.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 0: r = bB \ vA: out = out & "var.Byte0.idiv.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 0: r = vA Mod bB: out = out & "var.vDbl.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 0: r = bB Mod vA: out = out & "var.Byte0.mod.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 0: r = vA ^ bB: out = out & "var.vDbl.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 0: r = bB ^ vA: out = out & "var.Byte0.pow.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 200: r = vA + bB: out = out & "var.vDbl.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 200: r = bB + vA: out = out & "var.Byte1.add.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 200: r = vA - bB: out = out & "var.vDbl.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 200: r = bB - vA: out = out & "var.Byte1.sub.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 200: r = vA * bB: out = out & "var.vDbl.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 200: r = bB * vA: out = out & "var.Byte1.mul.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 200: r = vA / bB: out = out & "var.vDbl.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 200: r = bB / vA: out = out & "var.Byte1.div.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 200: r = vA \ bB: out = out & "var.vDbl.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 200: r = bB \ vA: out = out & "var.Byte1.idiv.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 200: r = vA Mod bB: out = out & "var.vDbl.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 200: r = bB Mod vA: out = out & "var.Byte1.mod.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 200: r = vA ^ bB: out = out & "var.vDbl.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 200: r = bB ^ vA: out = out & "var.Byte1.pow.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 255: r = vA + bB: out = out & "var.vDbl.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 255: r = bB + vA: out = out & "var.Byte2.add.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 255: r = vA - bB: out = out & "var.vDbl.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 255: r = bB - vA: out = out & "var.Byte2.sub.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 255: r = vA * bB: out = out & "var.vDbl.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 255: r = bB * vA: out = out & "var.Byte2.mul.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 255: r = vA / bB: out = out & "var.vDbl.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 255: r = bB / vA: out = out & "var.Byte2.div.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 255: r = vA \ bB: out = out & "var.vDbl.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 255: r = bB \ vA: out = out & "var.Byte2.idiv.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 255: r = vA Mod bB: out = out & "var.vDbl.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 255: r = bB Mod vA: out = out & "var.Byte2.mod.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 255: r = vA ^ bB: out = out & "var.vDbl.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: bB = 255: r = bB ^ vA: out = out & "var.Byte2.pow.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = -32768: r = vA + iB: out = out & "var.vDbl.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = -32768: r = iB + vA: out = out & "var.Integer0.add.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = -32768: r = vA - iB: out = out & "var.vDbl.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = -32768: r = iB - vA: out = out & "var.Integer0.sub.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = -32768: r = vA * iB: out = out & "var.vDbl.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = -32768: r = iB * vA: out = out & "var.Integer0.mul.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = -32768: r = vA / iB: out = out & "var.vDbl.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = -32768: r = iB / vA: out = out & "var.Integer0.div.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = -32768: r = vA \ iB: out = out & "var.vDbl.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = -32768: r = iB \ vA: out = out & "var.Integer0.idiv.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = -32768: r = vA Mod iB: out = out & "var.vDbl.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = -32768: r = iB Mod vA: out = out & "var.Integer0.mod.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = -32768: r = vA ^ iB: out = out & "var.vDbl.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = -32768: r = iB ^ vA: out = out & "var.Integer0.pow.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = 1: r = vA + iB: out = out & "var.vDbl.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = 1: r = iB + vA: out = out & "var.Integer1.add.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = 1: r = vA - iB: out = out & "var.vDbl.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = 1: r = iB - vA: out = out & "var.Integer1.sub.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = 1: r = vA * iB: out = out & "var.vDbl.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = 1: r = iB * vA: out = out & "var.Integer1.mul.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = 1: r = vA / iB: out = out & "var.vDbl.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = 1: r = iB / vA: out = out & "var.Integer1.div.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = 1: r = vA \ iB: out = out & "var.vDbl.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = 1: r = iB \ vA: out = out & "var.Integer1.idiv.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = 1: r = vA Mod iB: out = out & "var.vDbl.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  P038 = out
End Function
Function P039() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: vA = 2.5: iB = 1: r = iB Mod vA: out = out & "var.Integer1.mod.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = 1: r = vA ^ iB: out = out & "var.vDbl.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = 1: r = iB ^ vA: out = out & "var.Integer1.pow.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = 32767: r = vA + iB: out = out & "var.vDbl.add.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = 32767: r = iB + vA: out = out & "var.Integer2.add.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = 32767: r = vA - iB: out = out & "var.vDbl.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = 32767: r = iB - vA: out = out & "var.Integer2.sub.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = 32767: r = vA * iB: out = out & "var.vDbl.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = 32767: r = iB * vA: out = out & "var.Integer2.mul.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = 32767: r = vA / iB: out = out & "var.vDbl.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = 32767: r = iB / vA: out = out & "var.Integer2.div.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = 32767: r = vA \ iB: out = out & "var.vDbl.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = 32767: r = iB \ vA: out = out & "var.Integer2.idiv.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = 32767: r = vA Mod iB: out = out & "var.vDbl.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = 32767: r = iB Mod vA: out = out & "var.Integer2.mod.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = 32767: r = vA ^ iB: out = out & "var.vDbl.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: iB = 32767: r = iB ^ vA: out = out & "var.Integer2.pow.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = -2147483648: r = vA + lB: out = out & "var.vDbl.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = -2147483648: r = lB + vA: out = out & "var.Long0.add.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = -2147483648: r = vA - lB: out = out & "var.vDbl.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = -2147483648: r = lB - vA: out = out & "var.Long0.sub.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = -2147483648: r = vA * lB: out = out & "var.vDbl.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = -2147483648: r = lB * vA: out = out & "var.Long0.mul.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = -2147483648: r = vA / lB: out = out & "var.vDbl.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = -2147483648: r = lB / vA: out = out & "var.Long0.div.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = -2147483648: r = vA \ lB: out = out & "var.vDbl.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = -2147483648: r = lB \ vA: out = out & "var.Long0.idiv.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = -2147483648: r = vA Mod lB: out = out & "var.vDbl.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = -2147483648: r = lB Mod vA: out = out & "var.Long0.mod.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = -2147483648: r = vA ^ lB: out = out & "var.vDbl.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = -2147483648: r = lB ^ vA: out = out & "var.Long0.pow.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = 65536: r = vA + lB: out = out & "var.vDbl.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = 65536: r = lB + vA: out = out & "var.Long1.add.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = 65536: r = vA - lB: out = out & "var.vDbl.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = 65536: r = lB - vA: out = out & "var.Long1.sub.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = 65536: r = vA * lB: out = out & "var.vDbl.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = 65536: r = lB * vA: out = out & "var.Long1.mul.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = 65536: r = vA / lB: out = out & "var.vDbl.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = 65536: r = lB / vA: out = out & "var.Long1.div.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = 65536: r = vA \ lB: out = out & "var.vDbl.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = 65536: r = lB \ vA: out = out & "var.Long1.idiv.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = 65536: r = vA Mod lB: out = out & "var.vDbl.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = 65536: r = lB Mod vA: out = out & "var.Long1.mod.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = 65536: r = vA ^ lB: out = out & "var.vDbl.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = 65536: r = lB ^ vA: out = out & "var.Long1.pow.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = 2147483647: r = vA + lB: out = out & "var.vDbl.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = 2147483647: r = lB + vA: out = out & "var.Long2.add.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = 2147483647: r = vA - lB: out = out & "var.vDbl.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = 2147483647: r = lB - vA: out = out & "var.Long2.sub.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = 2147483647: r = vA * lB: out = out & "var.vDbl.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = 2147483647: r = lB * vA: out = out & "var.Long2.mul.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = 2147483647: r = vA / lB: out = out & "var.vDbl.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = 2147483647: r = lB / vA: out = out & "var.Long2.div.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = 2147483647: r = vA \ lB: out = out & "var.vDbl.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = 2147483647: r = lB \ vA: out = out & "var.Long2.idiv.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = 2147483647: r = vA Mod lB: out = out & "var.vDbl.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = 2147483647: r = lB Mod vA: out = out & "var.Long2.mod.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = 2147483647: r = vA ^ lB: out = out & "var.vDbl.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: lB = 2147483647: r = lB ^ vA: out = out & "var.Long2.pow.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = -9223372036854775807^: r = vA + qB: out = out & "var.vDbl.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = -9223372036854775807^: r = qB + vA: out = out & "var.LongLong0.add.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = -9223372036854775807^: r = vA - qB: out = out & "var.vDbl.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = -9223372036854775807^: r = qB - vA: out = out & "var.LongLong0.sub.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = -9223372036854775807^: r = vA * qB: out = out & "var.vDbl.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = -9223372036854775807^: r = qB * vA: out = out & "var.LongLong0.mul.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = -9223372036854775807^: r = vA / qB: out = out & "var.vDbl.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = -9223372036854775807^: r = qB / vA: out = out & "var.LongLong0.div.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = -9223372036854775807^: r = vA \ qB: out = out & "var.vDbl.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = -9223372036854775807^: r = qB \ vA: out = out & "var.LongLong0.idiv.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = -9223372036854775807^: r = vA Mod qB: out = out & "var.vDbl.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = -9223372036854775807^: r = qB Mod vA: out = out & "var.LongLong0.mod.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = -9223372036854775807^: r = vA ^ qB: out = out & "var.vDbl.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = -9223372036854775807^: r = qB ^ vA: out = out & "var.LongLong0.pow.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = 3: r = vA + qB: out = out & "var.vDbl.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = 3: r = qB + vA: out = out & "var.LongLong1.add.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = 3: r = vA - qB: out = out & "var.vDbl.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = 3: r = qB - vA: out = out & "var.LongLong1.sub.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = 3: r = vA * qB: out = out & "var.vDbl.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = 3: r = qB * vA: out = out & "var.LongLong1.mul.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = 3: r = vA / qB: out = out & "var.vDbl.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = 3: r = qB / vA: out = out & "var.LongLong1.div.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = 3: r = vA \ qB: out = out & "var.vDbl.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = 3: r = qB \ vA: out = out & "var.LongLong1.idiv.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = 3: r = vA Mod qB: out = out & "var.vDbl.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = 3: r = qB Mod vA: out = out & "var.LongLong1.mod.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = 3: r = vA ^ qB: out = out & "var.vDbl.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = 3: r = qB ^ vA: out = out & "var.LongLong1.pow.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = 9223372036854775807^: r = vA + qB: out = out & "var.vDbl.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = 9223372036854775807^: r = qB + vA: out = out & "var.LongLong2.add.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = 9223372036854775807^: r = vA - qB: out = out & "var.vDbl.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = 9223372036854775807^: r = qB - vA: out = out & "var.LongLong2.sub.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = 9223372036854775807^: r = vA * qB: out = out & "var.vDbl.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = 9223372036854775807^: r = qB * vA: out = out & "var.LongLong2.mul.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = 9223372036854775807^: r = vA / qB: out = out & "var.vDbl.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = 9223372036854775807^: r = qB / vA: out = out & "var.LongLong2.div.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = 9223372036854775807^: r = vA \ qB: out = out & "var.vDbl.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = 9223372036854775807^: r = qB \ vA: out = out & "var.LongLong2.idiv.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = 9223372036854775807^: r = vA Mod qB: out = out & "var.vDbl.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = 9223372036854775807^: r = qB Mod vA: out = out & "var.LongLong2.mod.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = 9223372036854775807^: r = vA ^ qB: out = out & "var.vDbl.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: qB = 9223372036854775807^: r = qB ^ vA: out = out & "var.LongLong2.pow.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 0.5: r = vA + sB: out = out & "var.vDbl.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 0.5: r = sB + vA: out = out & "var.Single0.add.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 0.5: r = vA - sB: out = out & "var.vDbl.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 0.5: r = sB - vA: out = out & "var.Single0.sub.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 0.5: r = vA * sB: out = out & "var.vDbl.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 0.5: r = sB * vA: out = out & "var.Single0.mul.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 0.5: r = vA / sB: out = out & "var.vDbl.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 0.5: r = sB / vA: out = out & "var.Single0.div.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 0.5: r = vA \ sB: out = out & "var.vDbl.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 0.5: r = sB \ vA: out = out & "var.Single0.idiv.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 0.5: r = vA Mod sB: out = out & "var.vDbl.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 0.5: r = sB Mod vA: out = out & "var.Single0.mod.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 0.5: r = vA ^ sB: out = out & "var.vDbl.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 0.5: r = sB ^ vA: out = out & "var.Single0.pow.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 16777216: r = vA + sB: out = out & "var.vDbl.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 16777216: r = sB + vA: out = out & "var.Single1.add.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 16777216: r = vA - sB: out = out & "var.vDbl.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 16777216: r = sB - vA: out = out & "var.Single1.sub.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 16777216: r = vA * sB: out = out & "var.vDbl.mul.Single1=" & Rz(Err.Number, r) & vbLf
  P039 = out
End Function
Function P040() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: vA = 2.5: sB = 16777216: r = sB * vA: out = out & "var.Single1.mul.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 16777216: r = vA / sB: out = out & "var.vDbl.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 16777216: r = sB / vA: out = out & "var.Single1.div.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 16777216: r = vA \ sB: out = out & "var.vDbl.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 16777216: r = sB \ vA: out = out & "var.Single1.idiv.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 16777216: r = vA Mod sB: out = out & "var.vDbl.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 16777216: r = sB Mod vA: out = out & "var.Single1.mod.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 16777216: r = vA ^ sB: out = out & "var.vDbl.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 16777216: r = sB ^ vA: out = out & "var.Single1.pow.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 3.0E+38: r = vA + sB: out = out & "var.vDbl.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 3.0E+38: r = sB + vA: out = out & "var.Single2.add.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 3.0E+38: r = vA - sB: out = out & "var.vDbl.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 3.0E+38: r = sB - vA: out = out & "var.Single2.sub.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 3.0E+38: r = vA * sB: out = out & "var.vDbl.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 3.0E+38: r = sB * vA: out = out & "var.Single2.mul.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 3.0E+38: r = vA / sB: out = out & "var.vDbl.div.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 3.0E+38: r = sB / vA: out = out & "var.Single2.div.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 3.0E+38: r = vA \ sB: out = out & "var.vDbl.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 3.0E+38: r = sB \ vA: out = out & "var.Single2.idiv.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 3.0E+38: r = vA Mod sB: out = out & "var.vDbl.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 3.0E+38: r = sB Mod vA: out = out & "var.Single2.mod.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 3.0E+38: r = vA ^ sB: out = out & "var.vDbl.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: sB = 3.0E+38: r = sB ^ vA: out = out & "var.Single2.pow.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 0.5: r = vA + dB: out = out & "var.vDbl.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 0.5: r = dB + vA: out = out & "var.Double0.add.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 0.5: r = vA - dB: out = out & "var.vDbl.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 0.5: r = dB - vA: out = out & "var.Double0.sub.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 0.5: r = vA * dB: out = out & "var.vDbl.mul.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 0.5: r = dB * vA: out = out & "var.Double0.mul.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 0.5: r = vA / dB: out = out & "var.vDbl.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 0.5: r = dB / vA: out = out & "var.Double0.div.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 0.5: r = vA \ dB: out = out & "var.vDbl.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 0.5: r = dB \ vA: out = out & "var.Double0.idiv.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 0.5: r = vA Mod dB: out = out & "var.vDbl.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 0.5: r = dB Mod vA: out = out & "var.Double0.mod.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 0.5: r = vA ^ dB: out = out & "var.vDbl.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 0.5: r = dB ^ vA: out = out & "var.Double0.pow.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 2.5: r = vA + dB: out = out & "var.vDbl.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 2.5: r = dB + vA: out = out & "var.Double1.add.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 2.5: r = vA - dB: out = out & "var.vDbl.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 2.5: r = dB - vA: out = out & "var.Double1.sub.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 2.5: r = vA * dB: out = out & "var.vDbl.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 2.5: r = dB * vA: out = out & "var.Double1.mul.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 2.5: r = vA / dB: out = out & "var.vDbl.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 2.5: r = dB / vA: out = out & "var.Double1.div.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 2.5: r = vA \ dB: out = out & "var.vDbl.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 2.5: r = dB \ vA: out = out & "var.Double1.idiv.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 2.5: r = vA Mod dB: out = out & "var.vDbl.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 2.5: r = dB Mod vA: out = out & "var.Double1.mod.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 2.5: r = vA ^ dB: out = out & "var.vDbl.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 2.5: r = dB ^ vA: out = out & "var.Double1.pow.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 1.0E+308: r = vA + dB: out = out & "var.vDbl.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 1.0E+308: r = dB + vA: out = out & "var.Double2.add.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 1.0E+308: r = vA - dB: out = out & "var.vDbl.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 1.0E+308: r = dB - vA: out = out & "var.Double2.sub.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 1.0E+308: r = vA * dB: out = out & "var.vDbl.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 1.0E+308: r = dB * vA: out = out & "var.Double2.mul.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 1.0E+308: r = vA / dB: out = out & "var.vDbl.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 1.0E+308: r = dB / vA: out = out & "var.Double2.div.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 1.0E+308: r = vA \ dB: out = out & "var.vDbl.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 1.0E+308: r = dB \ vA: out = out & "var.Double2.idiv.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 1.0E+308: r = vA Mod dB: out = out & "var.vDbl.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 1.0E+308: r = dB Mod vA: out = out & "var.Double2.mod.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 1.0E+308: r = vA ^ dB: out = out & "var.vDbl.pow.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: dB = 1.0E+308: r = dB ^ vA: out = out & "var.Double2.pow.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = -0.5@: r = vA + cB: out = out & "var.vDbl.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = -0.5@: r = cB + vA: out = out & "var.Currency0.add.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = -0.5@: r = vA - cB: out = out & "var.vDbl.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = -0.5@: r = cB - vA: out = out & "var.Currency0.sub.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = -0.5@: r = vA * cB: out = out & "var.vDbl.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = -0.5@: r = cB * vA: out = out & "var.Currency0.mul.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = -0.5@: r = vA / cB: out = out & "var.vDbl.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = -0.5@: r = cB / vA: out = out & "var.Currency0.div.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = -0.5@: r = vA \ cB: out = out & "var.vDbl.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = -0.5@: r = cB \ vA: out = out & "var.Currency0.idiv.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = -0.5@: r = vA Mod cB: out = out & "var.vDbl.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = -0.5@: r = cB Mod vA: out = out & "var.Currency0.mod.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = -0.5@: r = vA ^ cB: out = out & "var.vDbl.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = -0.5@: r = cB ^ vA: out = out & "var.Currency0.pow.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = 1.5@: r = vA + cB: out = out & "var.vDbl.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = 1.5@: r = cB + vA: out = out & "var.Currency1.add.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = 1.5@: r = vA - cB: out = out & "var.vDbl.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = 1.5@: r = cB - vA: out = out & "var.Currency1.sub.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = 1.5@: r = vA * cB: out = out & "var.vDbl.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = 1.5@: r = cB * vA: out = out & "var.Currency1.mul.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = 1.5@: r = vA / cB: out = out & "var.vDbl.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = 1.5@: r = cB / vA: out = out & "var.Currency1.div.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = 1.5@: r = vA \ cB: out = out & "var.vDbl.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = 1.5@: r = cB \ vA: out = out & "var.Currency1.idiv.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = 1.5@: r = vA Mod cB: out = out & "var.vDbl.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = 1.5@: r = cB Mod vA: out = out & "var.Currency1.mod.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = 1.5@: r = vA ^ cB: out = out & "var.vDbl.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = 1.5@: r = cB ^ vA: out = out & "var.Currency1.pow.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = 922337203685477.5807@: r = vA + cB: out = out & "var.vDbl.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = 922337203685477.5807@: r = cB + vA: out = out & "var.Currency2.add.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = 922337203685477.5807@: r = vA - cB: out = out & "var.vDbl.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = 922337203685477.5807@: r = cB - vA: out = out & "var.Currency2.sub.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = 922337203685477.5807@: r = vA * cB: out = out & "var.vDbl.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = 922337203685477.5807@: r = cB * vA: out = out & "var.Currency2.mul.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = 922337203685477.5807@: r = vA / cB: out = out & "var.vDbl.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = 922337203685477.5807@: r = cB / vA: out = out & "var.Currency2.div.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = 922337203685477.5807@: r = vA \ cB: out = out & "var.vDbl.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = 922337203685477.5807@: r = cB \ vA: out = out & "var.Currency2.idiv.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = 922337203685477.5807@: r = vA Mod cB: out = out & "var.vDbl.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = 922337203685477.5807@: r = cB Mod vA: out = out & "var.Currency2.mod.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = 922337203685477.5807@: r = vA ^ cB: out = out & "var.vDbl.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: cB = 922337203685477.5807@: r = cB ^ vA: out = out & "var.Currency2.pow.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 0: r = vA + bB: out = out & "var.vCur.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 0: r = bB + vA: out = out & "var.Byte0.add.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 0: r = vA - bB: out = out & "var.vCur.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 0: r = bB - vA: out = out & "var.Byte0.sub.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 0: r = vA * bB: out = out & "var.vCur.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 0: r = bB * vA: out = out & "var.Byte0.mul.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 0: r = vA / bB: out = out & "var.vCur.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 0: r = bB / vA: out = out & "var.Byte0.div.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 0: r = vA \ bB: out = out & "var.vCur.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 0: r = bB \ vA: out = out & "var.Byte0.idiv.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 0: r = vA Mod bB: out = out & "var.vCur.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 0: r = bB Mod vA: out = out & "var.Byte0.mod.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 0: r = vA ^ bB: out = out & "var.vCur.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  P040 = out
End Function
Function P041() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: vA = 1.5@: bB = 0: r = bB ^ vA: out = out & "var.Byte0.pow.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 200: r = vA + bB: out = out & "var.vCur.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 200: r = bB + vA: out = out & "var.Byte1.add.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 200: r = vA - bB: out = out & "var.vCur.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 200: r = bB - vA: out = out & "var.Byte1.sub.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 200: r = vA * bB: out = out & "var.vCur.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 200: r = bB * vA: out = out & "var.Byte1.mul.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 200: r = vA / bB: out = out & "var.vCur.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 200: r = bB / vA: out = out & "var.Byte1.div.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 200: r = vA \ bB: out = out & "var.vCur.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 200: r = bB \ vA: out = out & "var.Byte1.idiv.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 200: r = vA Mod bB: out = out & "var.vCur.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 200: r = bB Mod vA: out = out & "var.Byte1.mod.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 200: r = vA ^ bB: out = out & "var.vCur.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 200: r = bB ^ vA: out = out & "var.Byte1.pow.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 255: r = vA + bB: out = out & "var.vCur.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 255: r = bB + vA: out = out & "var.Byte2.add.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 255: r = vA - bB: out = out & "var.vCur.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 255: r = bB - vA: out = out & "var.Byte2.sub.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 255: r = vA * bB: out = out & "var.vCur.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 255: r = bB * vA: out = out & "var.Byte2.mul.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 255: r = vA / bB: out = out & "var.vCur.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 255: r = bB / vA: out = out & "var.Byte2.div.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 255: r = vA \ bB: out = out & "var.vCur.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 255: r = bB \ vA: out = out & "var.Byte2.idiv.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 255: r = vA Mod bB: out = out & "var.vCur.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 255: r = bB Mod vA: out = out & "var.Byte2.mod.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 255: r = vA ^ bB: out = out & "var.vCur.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: bB = 255: r = bB ^ vA: out = out & "var.Byte2.pow.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = -32768: r = vA + iB: out = out & "var.vCur.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = -32768: r = iB + vA: out = out & "var.Integer0.add.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = -32768: r = vA - iB: out = out & "var.vCur.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = -32768: r = iB - vA: out = out & "var.Integer0.sub.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = -32768: r = vA * iB: out = out & "var.vCur.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = -32768: r = iB * vA: out = out & "var.Integer0.mul.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = -32768: r = vA / iB: out = out & "var.vCur.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = -32768: r = iB / vA: out = out & "var.Integer0.div.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = -32768: r = vA \ iB: out = out & "var.vCur.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = -32768: r = iB \ vA: out = out & "var.Integer0.idiv.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = -32768: r = vA Mod iB: out = out & "var.vCur.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = -32768: r = iB Mod vA: out = out & "var.Integer0.mod.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = -32768: r = vA ^ iB: out = out & "var.vCur.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = -32768: r = iB ^ vA: out = out & "var.Integer0.pow.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = 1: r = vA + iB: out = out & "var.vCur.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = 1: r = iB + vA: out = out & "var.Integer1.add.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = 1: r = vA - iB: out = out & "var.vCur.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = 1: r = iB - vA: out = out & "var.Integer1.sub.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = 1: r = vA * iB: out = out & "var.vCur.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = 1: r = iB * vA: out = out & "var.Integer1.mul.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = 1: r = vA / iB: out = out & "var.vCur.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = 1: r = iB / vA: out = out & "var.Integer1.div.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = 1: r = vA \ iB: out = out & "var.vCur.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = 1: r = iB \ vA: out = out & "var.Integer1.idiv.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = 1: r = vA Mod iB: out = out & "var.vCur.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = 1: r = iB Mod vA: out = out & "var.Integer1.mod.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = 1: r = vA ^ iB: out = out & "var.vCur.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = 1: r = iB ^ vA: out = out & "var.Integer1.pow.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = 32767: r = vA + iB: out = out & "var.vCur.add.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = 32767: r = iB + vA: out = out & "var.Integer2.add.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = 32767: r = vA - iB: out = out & "var.vCur.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = 32767: r = iB - vA: out = out & "var.Integer2.sub.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = 32767: r = vA * iB: out = out & "var.vCur.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = 32767: r = iB * vA: out = out & "var.Integer2.mul.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = 32767: r = vA / iB: out = out & "var.vCur.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = 32767: r = iB / vA: out = out & "var.Integer2.div.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = 32767: r = vA \ iB: out = out & "var.vCur.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = 32767: r = iB \ vA: out = out & "var.Integer2.idiv.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = 32767: r = vA Mod iB: out = out & "var.vCur.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = 32767: r = iB Mod vA: out = out & "var.Integer2.mod.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = 32767: r = vA ^ iB: out = out & "var.vCur.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: iB = 32767: r = iB ^ vA: out = out & "var.Integer2.pow.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = -2147483648: r = vA + lB: out = out & "var.vCur.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = -2147483648: r = lB + vA: out = out & "var.Long0.add.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = -2147483648: r = vA - lB: out = out & "var.vCur.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = -2147483648: r = lB - vA: out = out & "var.Long0.sub.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = -2147483648: r = vA * lB: out = out & "var.vCur.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = -2147483648: r = lB * vA: out = out & "var.Long0.mul.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = -2147483648: r = vA / lB: out = out & "var.vCur.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = -2147483648: r = lB / vA: out = out & "var.Long0.div.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = -2147483648: r = vA \ lB: out = out & "var.vCur.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = -2147483648: r = lB \ vA: out = out & "var.Long0.idiv.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = -2147483648: r = vA Mod lB: out = out & "var.vCur.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = -2147483648: r = lB Mod vA: out = out & "var.Long0.mod.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = -2147483648: r = vA ^ lB: out = out & "var.vCur.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = -2147483648: r = lB ^ vA: out = out & "var.Long0.pow.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = 65536: r = vA + lB: out = out & "var.vCur.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = 65536: r = lB + vA: out = out & "var.Long1.add.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = 65536: r = vA - lB: out = out & "var.vCur.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = 65536: r = lB - vA: out = out & "var.Long1.sub.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = 65536: r = vA * lB: out = out & "var.vCur.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = 65536: r = lB * vA: out = out & "var.Long1.mul.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = 65536: r = vA / lB: out = out & "var.vCur.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = 65536: r = lB / vA: out = out & "var.Long1.div.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = 65536: r = vA \ lB: out = out & "var.vCur.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = 65536: r = lB \ vA: out = out & "var.Long1.idiv.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = 65536: r = vA Mod lB: out = out & "var.vCur.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = 65536: r = lB Mod vA: out = out & "var.Long1.mod.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = 65536: r = vA ^ lB: out = out & "var.vCur.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = 65536: r = lB ^ vA: out = out & "var.Long1.pow.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = 2147483647: r = vA + lB: out = out & "var.vCur.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = 2147483647: r = lB + vA: out = out & "var.Long2.add.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = 2147483647: r = vA - lB: out = out & "var.vCur.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = 2147483647: r = lB - vA: out = out & "var.Long2.sub.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = 2147483647: r = vA * lB: out = out & "var.vCur.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = 2147483647: r = lB * vA: out = out & "var.Long2.mul.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = 2147483647: r = vA / lB: out = out & "var.vCur.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = 2147483647: r = lB / vA: out = out & "var.Long2.div.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = 2147483647: r = vA \ lB: out = out & "var.vCur.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = 2147483647: r = lB \ vA: out = out & "var.Long2.idiv.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = 2147483647: r = vA Mod lB: out = out & "var.vCur.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = 2147483647: r = lB Mod vA: out = out & "var.Long2.mod.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = 2147483647: r = vA ^ lB: out = out & "var.vCur.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: lB = 2147483647: r = lB ^ vA: out = out & "var.Long2.pow.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = -9223372036854775807^: r = vA + qB: out = out & "var.vCur.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = -9223372036854775807^: r = qB + vA: out = out & "var.LongLong0.add.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = -9223372036854775807^: r = vA - qB: out = out & "var.vCur.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = -9223372036854775807^: r = qB - vA: out = out & "var.LongLong0.sub.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = -9223372036854775807^: r = vA * qB: out = out & "var.vCur.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = -9223372036854775807^: r = qB * vA: out = out & "var.LongLong0.mul.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = -9223372036854775807^: r = vA / qB: out = out & "var.vCur.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  P041 = out
End Function
Function P042() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: vA = 1.5@: qB = -9223372036854775807^: r = qB / vA: out = out & "var.LongLong0.div.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = -9223372036854775807^: r = vA \ qB: out = out & "var.vCur.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = -9223372036854775807^: r = qB \ vA: out = out & "var.LongLong0.idiv.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = -9223372036854775807^: r = vA Mod qB: out = out & "var.vCur.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = -9223372036854775807^: r = qB Mod vA: out = out & "var.LongLong0.mod.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = -9223372036854775807^: r = vA ^ qB: out = out & "var.vCur.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = -9223372036854775807^: r = qB ^ vA: out = out & "var.LongLong0.pow.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = 3: r = vA + qB: out = out & "var.vCur.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = 3: r = qB + vA: out = out & "var.LongLong1.add.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = 3: r = vA - qB: out = out & "var.vCur.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = 3: r = qB - vA: out = out & "var.LongLong1.sub.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = 3: r = vA * qB: out = out & "var.vCur.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = 3: r = qB * vA: out = out & "var.LongLong1.mul.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = 3: r = vA / qB: out = out & "var.vCur.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = 3: r = qB / vA: out = out & "var.LongLong1.div.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = 3: r = vA \ qB: out = out & "var.vCur.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = 3: r = qB \ vA: out = out & "var.LongLong1.idiv.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = 3: r = vA Mod qB: out = out & "var.vCur.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = 3: r = qB Mod vA: out = out & "var.LongLong1.mod.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = 3: r = vA ^ qB: out = out & "var.vCur.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = 3: r = qB ^ vA: out = out & "var.LongLong1.pow.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = 9223372036854775807^: r = vA + qB: out = out & "var.vCur.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = 9223372036854775807^: r = qB + vA: out = out & "var.LongLong2.add.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = 9223372036854775807^: r = vA - qB: out = out & "var.vCur.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = 9223372036854775807^: r = qB - vA: out = out & "var.LongLong2.sub.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = 9223372036854775807^: r = vA * qB: out = out & "var.vCur.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = 9223372036854775807^: r = qB * vA: out = out & "var.LongLong2.mul.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = 9223372036854775807^: r = vA / qB: out = out & "var.vCur.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = 9223372036854775807^: r = qB / vA: out = out & "var.LongLong2.div.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = 9223372036854775807^: r = vA \ qB: out = out & "var.vCur.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = 9223372036854775807^: r = qB \ vA: out = out & "var.LongLong2.idiv.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = 9223372036854775807^: r = vA Mod qB: out = out & "var.vCur.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = 9223372036854775807^: r = qB Mod vA: out = out & "var.LongLong2.mod.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = 9223372036854775807^: r = vA ^ qB: out = out & "var.vCur.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: qB = 9223372036854775807^: r = qB ^ vA: out = out & "var.LongLong2.pow.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 0.5: r = vA + sB: out = out & "var.vCur.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 0.5: r = sB + vA: out = out & "var.Single0.add.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 0.5: r = vA - sB: out = out & "var.vCur.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 0.5: r = sB - vA: out = out & "var.Single0.sub.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 0.5: r = vA * sB: out = out & "var.vCur.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 0.5: r = sB * vA: out = out & "var.Single0.mul.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 0.5: r = vA / sB: out = out & "var.vCur.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 0.5: r = sB / vA: out = out & "var.Single0.div.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 0.5: r = vA \ sB: out = out & "var.vCur.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 0.5: r = sB \ vA: out = out & "var.Single0.idiv.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 0.5: r = vA Mod sB: out = out & "var.vCur.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 0.5: r = sB Mod vA: out = out & "var.Single0.mod.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 0.5: r = vA ^ sB: out = out & "var.vCur.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 0.5: r = sB ^ vA: out = out & "var.Single0.pow.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 16777216: r = vA + sB: out = out & "var.vCur.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 16777216: r = sB + vA: out = out & "var.Single1.add.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 16777216: r = vA - sB: out = out & "var.vCur.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 16777216: r = sB - vA: out = out & "var.Single1.sub.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 16777216: r = vA * sB: out = out & "var.vCur.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 16777216: r = sB * vA: out = out & "var.Single1.mul.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 16777216: r = vA / sB: out = out & "var.vCur.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 16777216: r = sB / vA: out = out & "var.Single1.div.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 16777216: r = vA \ sB: out = out & "var.vCur.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 16777216: r = sB \ vA: out = out & "var.Single1.idiv.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 16777216: r = vA Mod sB: out = out & "var.vCur.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 16777216: r = sB Mod vA: out = out & "var.Single1.mod.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 16777216: r = vA ^ sB: out = out & "var.vCur.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 16777216: r = sB ^ vA: out = out & "var.Single1.pow.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 3.0E+38: r = vA + sB: out = out & "var.vCur.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 3.0E+38: r = sB + vA: out = out & "var.Single2.add.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 3.0E+38: r = vA - sB: out = out & "var.vCur.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 3.0E+38: r = sB - vA: out = out & "var.Single2.sub.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 3.0E+38: r = vA * sB: out = out & "var.vCur.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 3.0E+38: r = sB * vA: out = out & "var.Single2.mul.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 3.0E+38: r = vA / sB: out = out & "var.vCur.div.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 3.0E+38: r = sB / vA: out = out & "var.Single2.div.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 3.0E+38: r = vA \ sB: out = out & "var.vCur.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 3.0E+38: r = sB \ vA: out = out & "var.Single2.idiv.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 3.0E+38: r = vA Mod sB: out = out & "var.vCur.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 3.0E+38: r = sB Mod vA: out = out & "var.Single2.mod.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 3.0E+38: r = vA ^ sB: out = out & "var.vCur.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: sB = 3.0E+38: r = sB ^ vA: out = out & "var.Single2.pow.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 0.5: r = vA + dB: out = out & "var.vCur.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 0.5: r = dB + vA: out = out & "var.Double0.add.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 0.5: r = vA - dB: out = out & "var.vCur.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 0.5: r = dB - vA: out = out & "var.Double0.sub.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 0.5: r = vA * dB: out = out & "var.vCur.mul.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 0.5: r = dB * vA: out = out & "var.Double0.mul.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 0.5: r = vA / dB: out = out & "var.vCur.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 0.5: r = dB / vA: out = out & "var.Double0.div.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 0.5: r = vA \ dB: out = out & "var.vCur.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 0.5: r = dB \ vA: out = out & "var.Double0.idiv.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 0.5: r = vA Mod dB: out = out & "var.vCur.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 0.5: r = dB Mod vA: out = out & "var.Double0.mod.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 0.5: r = vA ^ dB: out = out & "var.vCur.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 0.5: r = dB ^ vA: out = out & "var.Double0.pow.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 2.5: r = vA + dB: out = out & "var.vCur.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 2.5: r = dB + vA: out = out & "var.Double1.add.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 2.5: r = vA - dB: out = out & "var.vCur.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 2.5: r = dB - vA: out = out & "var.Double1.sub.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 2.5: r = vA * dB: out = out & "var.vCur.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 2.5: r = dB * vA: out = out & "var.Double1.mul.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 2.5: r = vA / dB: out = out & "var.vCur.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 2.5: r = dB / vA: out = out & "var.Double1.div.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 2.5: r = vA \ dB: out = out & "var.vCur.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 2.5: r = dB \ vA: out = out & "var.Double1.idiv.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 2.5: r = vA Mod dB: out = out & "var.vCur.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 2.5: r = dB Mod vA: out = out & "var.Double1.mod.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 2.5: r = vA ^ dB: out = out & "var.vCur.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 2.5: r = dB ^ vA: out = out & "var.Double1.pow.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 1.0E+308: r = vA + dB: out = out & "var.vCur.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 1.0E+308: r = dB + vA: out = out & "var.Double2.add.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 1.0E+308: r = vA - dB: out = out & "var.vCur.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 1.0E+308: r = dB - vA: out = out & "var.Double2.sub.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 1.0E+308: r = vA * dB: out = out & "var.vCur.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 1.0E+308: r = dB * vA: out = out & "var.Double2.mul.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 1.0E+308: r = vA / dB: out = out & "var.vCur.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 1.0E+308: r = dB / vA: out = out & "var.Double2.div.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 1.0E+308: r = vA \ dB: out = out & "var.vCur.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 1.0E+308: r = dB \ vA: out = out & "var.Double2.idiv.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 1.0E+308: r = vA Mod dB: out = out & "var.vCur.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 1.0E+308: r = dB Mod vA: out = out & "var.Double2.mod.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 1.0E+308: r = vA ^ dB: out = out & "var.vCur.pow.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: dB = 1.0E+308: r = dB ^ vA: out = out & "var.Double2.pow.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = -0.5@: r = vA + cB: out = out & "var.vCur.add.Currency0=" & Rz(Err.Number, r) & vbLf
  P042 = out
End Function
Function P043() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: vA = 1.5@: cB = -0.5@: r = cB + vA: out = out & "var.Currency0.add.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = -0.5@: r = vA - cB: out = out & "var.vCur.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = -0.5@: r = cB - vA: out = out & "var.Currency0.sub.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = -0.5@: r = vA * cB: out = out & "var.vCur.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = -0.5@: r = cB * vA: out = out & "var.Currency0.mul.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = -0.5@: r = vA / cB: out = out & "var.vCur.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = -0.5@: r = cB / vA: out = out & "var.Currency0.div.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = -0.5@: r = vA \ cB: out = out & "var.vCur.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = -0.5@: r = cB \ vA: out = out & "var.Currency0.idiv.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = -0.5@: r = vA Mod cB: out = out & "var.vCur.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = -0.5@: r = cB Mod vA: out = out & "var.Currency0.mod.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = -0.5@: r = vA ^ cB: out = out & "var.vCur.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = -0.5@: r = cB ^ vA: out = out & "var.Currency0.pow.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = 1.5@: r = vA + cB: out = out & "var.vCur.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = 1.5@: r = cB + vA: out = out & "var.Currency1.add.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = 1.5@: r = vA - cB: out = out & "var.vCur.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = 1.5@: r = cB - vA: out = out & "var.Currency1.sub.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = 1.5@: r = vA * cB: out = out & "var.vCur.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = 1.5@: r = cB * vA: out = out & "var.Currency1.mul.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = 1.5@: r = vA / cB: out = out & "var.vCur.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = 1.5@: r = cB / vA: out = out & "var.Currency1.div.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = 1.5@: r = vA \ cB: out = out & "var.vCur.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = 1.5@: r = cB \ vA: out = out & "var.Currency1.idiv.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = 1.5@: r = vA Mod cB: out = out & "var.vCur.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = 1.5@: r = cB Mod vA: out = out & "var.Currency1.mod.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = 1.5@: r = vA ^ cB: out = out & "var.vCur.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = 1.5@: r = cB ^ vA: out = out & "var.Currency1.pow.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = 922337203685477.5807@: r = vA + cB: out = out & "var.vCur.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = 922337203685477.5807@: r = cB + vA: out = out & "var.Currency2.add.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = 922337203685477.5807@: r = vA - cB: out = out & "var.vCur.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = 922337203685477.5807@: r = cB - vA: out = out & "var.Currency2.sub.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = 922337203685477.5807@: r = vA * cB: out = out & "var.vCur.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = 922337203685477.5807@: r = cB * vA: out = out & "var.Currency2.mul.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = 922337203685477.5807@: r = vA / cB: out = out & "var.vCur.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = 922337203685477.5807@: r = cB / vA: out = out & "var.Currency2.div.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = 922337203685477.5807@: r = vA \ cB: out = out & "var.vCur.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = 922337203685477.5807@: r = cB \ vA: out = out & "var.Currency2.idiv.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = 922337203685477.5807@: r = vA Mod cB: out = out & "var.vCur.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = 922337203685477.5807@: r = cB Mod vA: out = out & "var.Currency2.mod.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = 922337203685477.5807@: r = vA ^ cB: out = out & "var.vCur.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: cB = 922337203685477.5807@: r = cB ^ vA: out = out & "var.Currency2.pow.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 0: r = vA + bB: out = out & "var.vStr.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 0: r = bB + vA: out = out & "var.Byte0.add.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 0: r = vA - bB: out = out & "var.vStr.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 0: r = bB - vA: out = out & "var.Byte0.sub.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 0: r = vA * bB: out = out & "var.vStr.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 0: r = bB * vA: out = out & "var.Byte0.mul.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 0: r = vA / bB: out = out & "var.vStr.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 0: r = bB / vA: out = out & "var.Byte0.div.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 0: r = vA \ bB: out = out & "var.vStr.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 0: r = bB \ vA: out = out & "var.Byte0.idiv.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 0: r = vA Mod bB: out = out & "var.vStr.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 0: r = bB Mod vA: out = out & "var.Byte0.mod.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 0: r = vA ^ bB: out = out & "var.vStr.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 0: r = bB ^ vA: out = out & "var.Byte0.pow.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 200: r = vA + bB: out = out & "var.vStr.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 200: r = bB + vA: out = out & "var.Byte1.add.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 200: r = vA - bB: out = out & "var.vStr.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 200: r = bB - vA: out = out & "var.Byte1.sub.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 200: r = vA * bB: out = out & "var.vStr.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 200: r = bB * vA: out = out & "var.Byte1.mul.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 200: r = vA / bB: out = out & "var.vStr.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 200: r = bB / vA: out = out & "var.Byte1.div.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 200: r = vA \ bB: out = out & "var.vStr.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 200: r = bB \ vA: out = out & "var.Byte1.idiv.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 200: r = vA Mod bB: out = out & "var.vStr.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 200: r = bB Mod vA: out = out & "var.Byte1.mod.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 200: r = vA ^ bB: out = out & "var.vStr.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 200: r = bB ^ vA: out = out & "var.Byte1.pow.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 255: r = vA + bB: out = out & "var.vStr.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 255: r = bB + vA: out = out & "var.Byte2.add.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 255: r = vA - bB: out = out & "var.vStr.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 255: r = bB - vA: out = out & "var.Byte2.sub.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 255: r = vA * bB: out = out & "var.vStr.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 255: r = bB * vA: out = out & "var.Byte2.mul.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 255: r = vA / bB: out = out & "var.vStr.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 255: r = bB / vA: out = out & "var.Byte2.div.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 255: r = vA \ bB: out = out & "var.vStr.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 255: r = bB \ vA: out = out & "var.Byte2.idiv.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 255: r = vA Mod bB: out = out & "var.vStr.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 255: r = bB Mod vA: out = out & "var.Byte2.mod.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 255: r = vA ^ bB: out = out & "var.vStr.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": bB = 255: r = bB ^ vA: out = out & "var.Byte2.pow.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = -32768: r = vA + iB: out = out & "var.vStr.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = -32768: r = iB + vA: out = out & "var.Integer0.add.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = -32768: r = vA - iB: out = out & "var.vStr.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = -32768: r = iB - vA: out = out & "var.Integer0.sub.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = -32768: r = vA * iB: out = out & "var.vStr.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = -32768: r = iB * vA: out = out & "var.Integer0.mul.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = -32768: r = vA / iB: out = out & "var.vStr.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = -32768: r = iB / vA: out = out & "var.Integer0.div.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = -32768: r = vA \ iB: out = out & "var.vStr.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = -32768: r = iB \ vA: out = out & "var.Integer0.idiv.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = -32768: r = vA Mod iB: out = out & "var.vStr.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = -32768: r = iB Mod vA: out = out & "var.Integer0.mod.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = -32768: r = vA ^ iB: out = out & "var.vStr.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = -32768: r = iB ^ vA: out = out & "var.Integer0.pow.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = 1: r = vA + iB: out = out & "var.vStr.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = 1: r = iB + vA: out = out & "var.Integer1.add.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = 1: r = vA - iB: out = out & "var.vStr.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = 1: r = iB - vA: out = out & "var.Integer1.sub.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = 1: r = vA * iB: out = out & "var.vStr.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = 1: r = iB * vA: out = out & "var.Integer1.mul.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = 1: r = vA / iB: out = out & "var.vStr.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = 1: r = iB / vA: out = out & "var.Integer1.div.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = 1: r = vA \ iB: out = out & "var.vStr.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = 1: r = iB \ vA: out = out & "var.Integer1.idiv.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = 1: r = vA Mod iB: out = out & "var.vStr.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = 1: r = iB Mod vA: out = out & "var.Integer1.mod.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = 1: r = vA ^ iB: out = out & "var.vStr.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = 1: r = iB ^ vA: out = out & "var.Integer1.pow.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = 32767: r = vA + iB: out = out & "var.vStr.add.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = 32767: r = iB + vA: out = out & "var.Integer2.add.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = 32767: r = vA - iB: out = out & "var.vStr.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = 32767: r = iB - vA: out = out & "var.Integer2.sub.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = 32767: r = vA * iB: out = out & "var.vStr.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = 32767: r = iB * vA: out = out & "var.Integer2.mul.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = 32767: r = vA / iB: out = out & "var.vStr.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = 32767: r = iB / vA: out = out & "var.Integer2.div.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = 32767: r = vA \ iB: out = out & "var.vStr.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  P043 = out
End Function
Function P044() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: vA = "12": iB = 32767: r = iB \ vA: out = out & "var.Integer2.idiv.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = 32767: r = vA Mod iB: out = out & "var.vStr.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = 32767: r = iB Mod vA: out = out & "var.Integer2.mod.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = 32767: r = vA ^ iB: out = out & "var.vStr.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": iB = 32767: r = iB ^ vA: out = out & "var.Integer2.pow.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = -2147483648: r = vA + lB: out = out & "var.vStr.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = -2147483648: r = lB + vA: out = out & "var.Long0.add.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = -2147483648: r = vA - lB: out = out & "var.vStr.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = -2147483648: r = lB - vA: out = out & "var.Long0.sub.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = -2147483648: r = vA * lB: out = out & "var.vStr.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = -2147483648: r = lB * vA: out = out & "var.Long0.mul.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = -2147483648: r = vA / lB: out = out & "var.vStr.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = -2147483648: r = lB / vA: out = out & "var.Long0.div.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = -2147483648: r = vA \ lB: out = out & "var.vStr.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = -2147483648: r = lB \ vA: out = out & "var.Long0.idiv.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = -2147483648: r = vA Mod lB: out = out & "var.vStr.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = -2147483648: r = lB Mod vA: out = out & "var.Long0.mod.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = -2147483648: r = vA ^ lB: out = out & "var.vStr.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = -2147483648: r = lB ^ vA: out = out & "var.Long0.pow.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = 65536: r = vA + lB: out = out & "var.vStr.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = 65536: r = lB + vA: out = out & "var.Long1.add.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = 65536: r = vA - lB: out = out & "var.vStr.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = 65536: r = lB - vA: out = out & "var.Long1.sub.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = 65536: r = vA * lB: out = out & "var.vStr.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = 65536: r = lB * vA: out = out & "var.Long1.mul.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = 65536: r = vA / lB: out = out & "var.vStr.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = 65536: r = lB / vA: out = out & "var.Long1.div.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = 65536: r = vA \ lB: out = out & "var.vStr.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = 65536: r = lB \ vA: out = out & "var.Long1.idiv.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = 65536: r = vA Mod lB: out = out & "var.vStr.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = 65536: r = lB Mod vA: out = out & "var.Long1.mod.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = 65536: r = vA ^ lB: out = out & "var.vStr.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = 65536: r = lB ^ vA: out = out & "var.Long1.pow.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = 2147483647: r = vA + lB: out = out & "var.vStr.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = 2147483647: r = lB + vA: out = out & "var.Long2.add.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = 2147483647: r = vA - lB: out = out & "var.vStr.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = 2147483647: r = lB - vA: out = out & "var.Long2.sub.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = 2147483647: r = vA * lB: out = out & "var.vStr.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = 2147483647: r = lB * vA: out = out & "var.Long2.mul.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = 2147483647: r = vA / lB: out = out & "var.vStr.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = 2147483647: r = lB / vA: out = out & "var.Long2.div.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = 2147483647: r = vA \ lB: out = out & "var.vStr.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = 2147483647: r = lB \ vA: out = out & "var.Long2.idiv.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = 2147483647: r = vA Mod lB: out = out & "var.vStr.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = 2147483647: r = lB Mod vA: out = out & "var.Long2.mod.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = 2147483647: r = vA ^ lB: out = out & "var.vStr.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": lB = 2147483647: r = lB ^ vA: out = out & "var.Long2.pow.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = -9223372036854775807^: r = vA + qB: out = out & "var.vStr.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = -9223372036854775807^: r = qB + vA: out = out & "var.LongLong0.add.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = -9223372036854775807^: r = vA - qB: out = out & "var.vStr.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = -9223372036854775807^: r = qB - vA: out = out & "var.LongLong0.sub.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = -9223372036854775807^: r = vA * qB: out = out & "var.vStr.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = -9223372036854775807^: r = qB * vA: out = out & "var.LongLong0.mul.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = -9223372036854775807^: r = vA / qB: out = out & "var.vStr.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = -9223372036854775807^: r = qB / vA: out = out & "var.LongLong0.div.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = -9223372036854775807^: r = vA \ qB: out = out & "var.vStr.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = -9223372036854775807^: r = qB \ vA: out = out & "var.LongLong0.idiv.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = -9223372036854775807^: r = vA Mod qB: out = out & "var.vStr.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = -9223372036854775807^: r = qB Mod vA: out = out & "var.LongLong0.mod.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = -9223372036854775807^: r = vA ^ qB: out = out & "var.vStr.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = -9223372036854775807^: r = qB ^ vA: out = out & "var.LongLong0.pow.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = 3: r = vA + qB: out = out & "var.vStr.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = 3: r = qB + vA: out = out & "var.LongLong1.add.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = 3: r = vA - qB: out = out & "var.vStr.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = 3: r = qB - vA: out = out & "var.LongLong1.sub.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = 3: r = vA * qB: out = out & "var.vStr.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = 3: r = qB * vA: out = out & "var.LongLong1.mul.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = 3: r = vA / qB: out = out & "var.vStr.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = 3: r = qB / vA: out = out & "var.LongLong1.div.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = 3: r = vA \ qB: out = out & "var.vStr.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = 3: r = qB \ vA: out = out & "var.LongLong1.idiv.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = 3: r = vA Mod qB: out = out & "var.vStr.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = 3: r = qB Mod vA: out = out & "var.LongLong1.mod.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = 3: r = vA ^ qB: out = out & "var.vStr.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = 3: r = qB ^ vA: out = out & "var.LongLong1.pow.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = 9223372036854775807^: r = vA + qB: out = out & "var.vStr.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = 9223372036854775807^: r = qB + vA: out = out & "var.LongLong2.add.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = 9223372036854775807^: r = vA - qB: out = out & "var.vStr.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = 9223372036854775807^: r = qB - vA: out = out & "var.LongLong2.sub.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = 9223372036854775807^: r = vA * qB: out = out & "var.vStr.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = 9223372036854775807^: r = qB * vA: out = out & "var.LongLong2.mul.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = 9223372036854775807^: r = vA / qB: out = out & "var.vStr.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = 9223372036854775807^: r = qB / vA: out = out & "var.LongLong2.div.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = 9223372036854775807^: r = vA \ qB: out = out & "var.vStr.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = 9223372036854775807^: r = qB \ vA: out = out & "var.LongLong2.idiv.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = 9223372036854775807^: r = vA Mod qB: out = out & "var.vStr.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = 9223372036854775807^: r = qB Mod vA: out = out & "var.LongLong2.mod.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = 9223372036854775807^: r = vA ^ qB: out = out & "var.vStr.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": qB = 9223372036854775807^: r = qB ^ vA: out = out & "var.LongLong2.pow.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 0.5: r = vA + sB: out = out & "var.vStr.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 0.5: r = sB + vA: out = out & "var.Single0.add.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 0.5: r = vA - sB: out = out & "var.vStr.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 0.5: r = sB - vA: out = out & "var.Single0.sub.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 0.5: r = vA * sB: out = out & "var.vStr.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 0.5: r = sB * vA: out = out & "var.Single0.mul.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 0.5: r = vA / sB: out = out & "var.vStr.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 0.5: r = sB / vA: out = out & "var.Single0.div.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 0.5: r = vA \ sB: out = out & "var.vStr.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 0.5: r = sB \ vA: out = out & "var.Single0.idiv.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 0.5: r = vA Mod sB: out = out & "var.vStr.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 0.5: r = sB Mod vA: out = out & "var.Single0.mod.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 0.5: r = vA ^ sB: out = out & "var.vStr.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 0.5: r = sB ^ vA: out = out & "var.Single0.pow.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 16777216: r = vA + sB: out = out & "var.vStr.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 16777216: r = sB + vA: out = out & "var.Single1.add.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 16777216: r = vA - sB: out = out & "var.vStr.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 16777216: r = sB - vA: out = out & "var.Single1.sub.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 16777216: r = vA * sB: out = out & "var.vStr.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 16777216: r = sB * vA: out = out & "var.Single1.mul.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 16777216: r = vA / sB: out = out & "var.vStr.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 16777216: r = sB / vA: out = out & "var.Single1.div.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 16777216: r = vA \ sB: out = out & "var.vStr.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 16777216: r = sB \ vA: out = out & "var.Single1.idiv.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 16777216: r = vA Mod sB: out = out & "var.vStr.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 16777216: r = sB Mod vA: out = out & "var.Single1.mod.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 16777216: r = vA ^ sB: out = out & "var.vStr.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 16777216: r = sB ^ vA: out = out & "var.Single1.pow.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 3.0E+38: r = vA + sB: out = out & "var.vStr.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 3.0E+38: r = sB + vA: out = out & "var.Single2.add.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 3.0E+38: r = vA - sB: out = out & "var.vStr.sub.Single2=" & Rz(Err.Number, r) & vbLf
  P044 = out
End Function
Function P045() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: vA = "12": sB = 3.0E+38: r = sB - vA: out = out & "var.Single2.sub.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 3.0E+38: r = vA * sB: out = out & "var.vStr.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 3.0E+38: r = sB * vA: out = out & "var.Single2.mul.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 3.0E+38: r = vA / sB: out = out & "var.vStr.div.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 3.0E+38: r = sB / vA: out = out & "var.Single2.div.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 3.0E+38: r = vA \ sB: out = out & "var.vStr.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 3.0E+38: r = sB \ vA: out = out & "var.Single2.idiv.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 3.0E+38: r = vA Mod sB: out = out & "var.vStr.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 3.0E+38: r = sB Mod vA: out = out & "var.Single2.mod.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 3.0E+38: r = vA ^ sB: out = out & "var.vStr.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": sB = 3.0E+38: r = sB ^ vA: out = out & "var.Single2.pow.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 0.5: r = vA + dB: out = out & "var.vStr.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 0.5: r = dB + vA: out = out & "var.Double0.add.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 0.5: r = vA - dB: out = out & "var.vStr.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 0.5: r = dB - vA: out = out & "var.Double0.sub.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 0.5: r = vA * dB: out = out & "var.vStr.mul.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 0.5: r = dB * vA: out = out & "var.Double0.mul.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 0.5: r = vA / dB: out = out & "var.vStr.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 0.5: r = dB / vA: out = out & "var.Double0.div.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 0.5: r = vA \ dB: out = out & "var.vStr.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 0.5: r = dB \ vA: out = out & "var.Double0.idiv.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 0.5: r = vA Mod dB: out = out & "var.vStr.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 0.5: r = dB Mod vA: out = out & "var.Double0.mod.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 0.5: r = vA ^ dB: out = out & "var.vStr.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 0.5: r = dB ^ vA: out = out & "var.Double0.pow.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 2.5: r = vA + dB: out = out & "var.vStr.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 2.5: r = dB + vA: out = out & "var.Double1.add.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 2.5: r = vA - dB: out = out & "var.vStr.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 2.5: r = dB - vA: out = out & "var.Double1.sub.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 2.5: r = vA * dB: out = out & "var.vStr.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 2.5: r = dB * vA: out = out & "var.Double1.mul.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 2.5: r = vA / dB: out = out & "var.vStr.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 2.5: r = dB / vA: out = out & "var.Double1.div.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 2.5: r = vA \ dB: out = out & "var.vStr.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 2.5: r = dB \ vA: out = out & "var.Double1.idiv.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 2.5: r = vA Mod dB: out = out & "var.vStr.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 2.5: r = dB Mod vA: out = out & "var.Double1.mod.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 2.5: r = vA ^ dB: out = out & "var.vStr.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 2.5: r = dB ^ vA: out = out & "var.Double1.pow.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 1.0E+308: r = vA + dB: out = out & "var.vStr.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 1.0E+308: r = dB + vA: out = out & "var.Double2.add.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 1.0E+308: r = vA - dB: out = out & "var.vStr.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 1.0E+308: r = dB - vA: out = out & "var.Double2.sub.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 1.0E+308: r = vA * dB: out = out & "var.vStr.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 1.0E+308: r = dB * vA: out = out & "var.Double2.mul.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 1.0E+308: r = vA / dB: out = out & "var.vStr.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 1.0E+308: r = dB / vA: out = out & "var.Double2.div.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 1.0E+308: r = vA \ dB: out = out & "var.vStr.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 1.0E+308: r = dB \ vA: out = out & "var.Double2.idiv.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 1.0E+308: r = vA Mod dB: out = out & "var.vStr.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 1.0E+308: r = dB Mod vA: out = out & "var.Double2.mod.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 1.0E+308: r = vA ^ dB: out = out & "var.vStr.pow.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": dB = 1.0E+308: r = dB ^ vA: out = out & "var.Double2.pow.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = -0.5@: r = vA + cB: out = out & "var.vStr.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = -0.5@: r = cB + vA: out = out & "var.Currency0.add.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = -0.5@: r = vA - cB: out = out & "var.vStr.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = -0.5@: r = cB - vA: out = out & "var.Currency0.sub.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = -0.5@: r = vA * cB: out = out & "var.vStr.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = -0.5@: r = cB * vA: out = out & "var.Currency0.mul.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = -0.5@: r = vA / cB: out = out & "var.vStr.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = -0.5@: r = cB / vA: out = out & "var.Currency0.div.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = -0.5@: r = vA \ cB: out = out & "var.vStr.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = -0.5@: r = cB \ vA: out = out & "var.Currency0.idiv.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = -0.5@: r = vA Mod cB: out = out & "var.vStr.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = -0.5@: r = cB Mod vA: out = out & "var.Currency0.mod.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = -0.5@: r = vA ^ cB: out = out & "var.vStr.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = -0.5@: r = cB ^ vA: out = out & "var.Currency0.pow.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = 1.5@: r = vA + cB: out = out & "var.vStr.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = 1.5@: r = cB + vA: out = out & "var.Currency1.add.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = 1.5@: r = vA - cB: out = out & "var.vStr.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = 1.5@: r = cB - vA: out = out & "var.Currency1.sub.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = 1.5@: r = vA * cB: out = out & "var.vStr.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = 1.5@: r = cB * vA: out = out & "var.Currency1.mul.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = 1.5@: r = vA / cB: out = out & "var.vStr.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = 1.5@: r = cB / vA: out = out & "var.Currency1.div.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = 1.5@: r = vA \ cB: out = out & "var.vStr.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = 1.5@: r = cB \ vA: out = out & "var.Currency1.idiv.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = 1.5@: r = vA Mod cB: out = out & "var.vStr.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = 1.5@: r = cB Mod vA: out = out & "var.Currency1.mod.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = 1.5@: r = vA ^ cB: out = out & "var.vStr.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = 1.5@: r = cB ^ vA: out = out & "var.Currency1.pow.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = 922337203685477.5807@: r = vA + cB: out = out & "var.vStr.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = 922337203685477.5807@: r = cB + vA: out = out & "var.Currency2.add.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = 922337203685477.5807@: r = vA - cB: out = out & "var.vStr.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = 922337203685477.5807@: r = cB - vA: out = out & "var.Currency2.sub.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = 922337203685477.5807@: r = vA * cB: out = out & "var.vStr.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = 922337203685477.5807@: r = cB * vA: out = out & "var.Currency2.mul.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = 922337203685477.5807@: r = vA / cB: out = out & "var.vStr.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = 922337203685477.5807@: r = cB / vA: out = out & "var.Currency2.div.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = 922337203685477.5807@: r = vA \ cB: out = out & "var.vStr.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = 922337203685477.5807@: r = cB \ vA: out = out & "var.Currency2.idiv.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = 922337203685477.5807@: r = vA Mod cB: out = out & "var.vStr.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = 922337203685477.5807@: r = cB Mod vA: out = out & "var.Currency2.mod.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = 922337203685477.5807@: r = vA ^ cB: out = out & "var.vStr.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": cB = 922337203685477.5807@: r = cB ^ vA: out = out & "var.Currency2.pow.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 0: r = vA + bB: out = out & "var.vEmp.add.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 0: r = bB + vA: out = out & "var.Byte0.add.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 0: r = vA - bB: out = out & "var.vEmp.sub.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 0: r = bB - vA: out = out & "var.Byte0.sub.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 0: r = vA * bB: out = out & "var.vEmp.mul.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 0: r = bB * vA: out = out & "var.Byte0.mul.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 0: r = vA / bB: out = out & "var.vEmp.div.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 0: r = bB / vA: out = out & "var.Byte0.div.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 0: r = vA \ bB: out = out & "var.vEmp.idiv.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 0: r = bB \ vA: out = out & "var.Byte0.idiv.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 0: r = vA Mod bB: out = out & "var.vEmp.mod.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 0: r = bB Mod vA: out = out & "var.Byte0.mod.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 0: r = vA ^ bB: out = out & "var.vEmp.pow.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 0: r = bB ^ vA: out = out & "var.Byte0.pow.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 200: r = vA + bB: out = out & "var.vEmp.add.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 200: r = bB + vA: out = out & "var.Byte1.add.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 200: r = vA - bB: out = out & "var.vEmp.sub.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 200: r = bB - vA: out = out & "var.Byte1.sub.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 200: r = vA * bB: out = out & "var.vEmp.mul.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 200: r = bB * vA: out = out & "var.Byte1.mul.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 200: r = vA / bB: out = out & "var.vEmp.div.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 200: r = bB / vA: out = out & "var.Byte1.div.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 200: r = vA \ bB: out = out & "var.vEmp.idiv.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 200: r = bB \ vA: out = out & "var.Byte1.idiv.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 200: r = vA Mod bB: out = out & "var.vEmp.mod.Byte1=" & Rz(Err.Number, r) & vbLf
  P045 = out
End Function
Function P046() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: vA = Empty: bB = 200: r = bB Mod vA: out = out & "var.Byte1.mod.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 200: r = vA ^ bB: out = out & "var.vEmp.pow.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 200: r = bB ^ vA: out = out & "var.Byte1.pow.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 255: r = vA + bB: out = out & "var.vEmp.add.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 255: r = bB + vA: out = out & "var.Byte2.add.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 255: r = vA - bB: out = out & "var.vEmp.sub.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 255: r = bB - vA: out = out & "var.Byte2.sub.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 255: r = vA * bB: out = out & "var.vEmp.mul.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 255: r = bB * vA: out = out & "var.Byte2.mul.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 255: r = vA / bB: out = out & "var.vEmp.div.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 255: r = bB / vA: out = out & "var.Byte2.div.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 255: r = vA \ bB: out = out & "var.vEmp.idiv.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 255: r = bB \ vA: out = out & "var.Byte2.idiv.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 255: r = vA Mod bB: out = out & "var.vEmp.mod.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 255: r = bB Mod vA: out = out & "var.Byte2.mod.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 255: r = vA ^ bB: out = out & "var.vEmp.pow.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: bB = 255: r = bB ^ vA: out = out & "var.Byte2.pow.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = -32768: r = vA + iB: out = out & "var.vEmp.add.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = -32768: r = iB + vA: out = out & "var.Integer0.add.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = -32768: r = vA - iB: out = out & "var.vEmp.sub.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = -32768: r = iB - vA: out = out & "var.Integer0.sub.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = -32768: r = vA * iB: out = out & "var.vEmp.mul.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = -32768: r = iB * vA: out = out & "var.Integer0.mul.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = -32768: r = vA / iB: out = out & "var.vEmp.div.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = -32768: r = iB / vA: out = out & "var.Integer0.div.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = -32768: r = vA \ iB: out = out & "var.vEmp.idiv.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = -32768: r = iB \ vA: out = out & "var.Integer0.idiv.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = -32768: r = vA Mod iB: out = out & "var.vEmp.mod.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = -32768: r = iB Mod vA: out = out & "var.Integer0.mod.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = -32768: r = vA ^ iB: out = out & "var.vEmp.pow.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = -32768: r = iB ^ vA: out = out & "var.Integer0.pow.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 1: r = vA + iB: out = out & "var.vEmp.add.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 1: r = iB + vA: out = out & "var.Integer1.add.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 1: r = vA - iB: out = out & "var.vEmp.sub.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 1: r = iB - vA: out = out & "var.Integer1.sub.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 1: r = vA * iB: out = out & "var.vEmp.mul.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 1: r = iB * vA: out = out & "var.Integer1.mul.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 1: r = vA / iB: out = out & "var.vEmp.div.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 1: r = iB / vA: out = out & "var.Integer1.div.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 1: r = vA \ iB: out = out & "var.vEmp.idiv.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 1: r = iB \ vA: out = out & "var.Integer1.idiv.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 1: r = vA Mod iB: out = out & "var.vEmp.mod.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 1: r = iB Mod vA: out = out & "var.Integer1.mod.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 1: r = vA ^ iB: out = out & "var.vEmp.pow.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 1: r = iB ^ vA: out = out & "var.Integer1.pow.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 32767: r = vA + iB: out = out & "var.vEmp.add.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 32767: r = iB + vA: out = out & "var.Integer2.add.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 32767: r = vA - iB: out = out & "var.vEmp.sub.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 32767: r = iB - vA: out = out & "var.Integer2.sub.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 32767: r = vA * iB: out = out & "var.vEmp.mul.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 32767: r = iB * vA: out = out & "var.Integer2.mul.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 32767: r = vA / iB: out = out & "var.vEmp.div.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 32767: r = iB / vA: out = out & "var.Integer2.div.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 32767: r = vA \ iB: out = out & "var.vEmp.idiv.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 32767: r = iB \ vA: out = out & "var.Integer2.idiv.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 32767: r = vA Mod iB: out = out & "var.vEmp.mod.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 32767: r = iB Mod vA: out = out & "var.Integer2.mod.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 32767: r = vA ^ iB: out = out & "var.vEmp.pow.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: iB = 32767: r = iB ^ vA: out = out & "var.Integer2.pow.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = -2147483648: r = vA + lB: out = out & "var.vEmp.add.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = -2147483648: r = lB + vA: out = out & "var.Long0.add.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = -2147483648: r = vA - lB: out = out & "var.vEmp.sub.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = -2147483648: r = lB - vA: out = out & "var.Long0.sub.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = -2147483648: r = vA * lB: out = out & "var.vEmp.mul.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = -2147483648: r = lB * vA: out = out & "var.Long0.mul.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = -2147483648: r = vA / lB: out = out & "var.vEmp.div.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = -2147483648: r = lB / vA: out = out & "var.Long0.div.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = -2147483648: r = vA \ lB: out = out & "var.vEmp.idiv.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = -2147483648: r = lB \ vA: out = out & "var.Long0.idiv.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = -2147483648: r = vA Mod lB: out = out & "var.vEmp.mod.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = -2147483648: r = lB Mod vA: out = out & "var.Long0.mod.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = -2147483648: r = vA ^ lB: out = out & "var.vEmp.pow.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = -2147483648: r = lB ^ vA: out = out & "var.Long0.pow.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 65536: r = vA + lB: out = out & "var.vEmp.add.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 65536: r = lB + vA: out = out & "var.Long1.add.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 65536: r = vA - lB: out = out & "var.vEmp.sub.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 65536: r = lB - vA: out = out & "var.Long1.sub.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 65536: r = vA * lB: out = out & "var.vEmp.mul.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 65536: r = lB * vA: out = out & "var.Long1.mul.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 65536: r = vA / lB: out = out & "var.vEmp.div.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 65536: r = lB / vA: out = out & "var.Long1.div.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 65536: r = vA \ lB: out = out & "var.vEmp.idiv.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 65536: r = lB \ vA: out = out & "var.Long1.idiv.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 65536: r = vA Mod lB: out = out & "var.vEmp.mod.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 65536: r = lB Mod vA: out = out & "var.Long1.mod.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 65536: r = vA ^ lB: out = out & "var.vEmp.pow.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 65536: r = lB ^ vA: out = out & "var.Long1.pow.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 2147483647: r = vA + lB: out = out & "var.vEmp.add.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 2147483647: r = lB + vA: out = out & "var.Long2.add.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 2147483647: r = vA - lB: out = out & "var.vEmp.sub.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 2147483647: r = lB - vA: out = out & "var.Long2.sub.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 2147483647: r = vA * lB: out = out & "var.vEmp.mul.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 2147483647: r = lB * vA: out = out & "var.Long2.mul.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 2147483647: r = vA / lB: out = out & "var.vEmp.div.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 2147483647: r = lB / vA: out = out & "var.Long2.div.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 2147483647: r = vA \ lB: out = out & "var.vEmp.idiv.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 2147483647: r = lB \ vA: out = out & "var.Long2.idiv.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 2147483647: r = vA Mod lB: out = out & "var.vEmp.mod.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 2147483647: r = lB Mod vA: out = out & "var.Long2.mod.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 2147483647: r = vA ^ lB: out = out & "var.vEmp.pow.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: lB = 2147483647: r = lB ^ vA: out = out & "var.Long2.pow.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = -9223372036854775807^: r = vA + qB: out = out & "var.vEmp.add.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = -9223372036854775807^: r = qB + vA: out = out & "var.LongLong0.add.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = -9223372036854775807^: r = vA - qB: out = out & "var.vEmp.sub.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = -9223372036854775807^: r = qB - vA: out = out & "var.LongLong0.sub.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = -9223372036854775807^: r = vA * qB: out = out & "var.vEmp.mul.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = -9223372036854775807^: r = qB * vA: out = out & "var.LongLong0.mul.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = -9223372036854775807^: r = vA / qB: out = out & "var.vEmp.div.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = -9223372036854775807^: r = qB / vA: out = out & "var.LongLong0.div.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = -9223372036854775807^: r = vA \ qB: out = out & "var.vEmp.idiv.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = -9223372036854775807^: r = qB \ vA: out = out & "var.LongLong0.idiv.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = -9223372036854775807^: r = vA Mod qB: out = out & "var.vEmp.mod.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = -9223372036854775807^: r = qB Mod vA: out = out & "var.LongLong0.mod.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = -9223372036854775807^: r = vA ^ qB: out = out & "var.vEmp.pow.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = -9223372036854775807^: r = qB ^ vA: out = out & "var.LongLong0.pow.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 3: r = vA + qB: out = out & "var.vEmp.add.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 3: r = qB + vA: out = out & "var.LongLong1.add.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 3: r = vA - qB: out = out & "var.vEmp.sub.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 3: r = qB - vA: out = out & "var.LongLong1.sub.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 3: r = vA * qB: out = out & "var.vEmp.mul.LongLong1=" & Rz(Err.Number, r) & vbLf
  P046 = out
End Function
Function P047() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: vA = Empty: qB = 3: r = qB * vA: out = out & "var.LongLong1.mul.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 3: r = vA / qB: out = out & "var.vEmp.div.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 3: r = qB / vA: out = out & "var.LongLong1.div.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 3: r = vA \ qB: out = out & "var.vEmp.idiv.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 3: r = qB \ vA: out = out & "var.LongLong1.idiv.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 3: r = vA Mod qB: out = out & "var.vEmp.mod.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 3: r = qB Mod vA: out = out & "var.LongLong1.mod.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 3: r = vA ^ qB: out = out & "var.vEmp.pow.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 3: r = qB ^ vA: out = out & "var.LongLong1.pow.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 9223372036854775807^: r = vA + qB: out = out & "var.vEmp.add.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 9223372036854775807^: r = qB + vA: out = out & "var.LongLong2.add.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 9223372036854775807^: r = vA - qB: out = out & "var.vEmp.sub.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 9223372036854775807^: r = qB - vA: out = out & "var.LongLong2.sub.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 9223372036854775807^: r = vA * qB: out = out & "var.vEmp.mul.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 9223372036854775807^: r = qB * vA: out = out & "var.LongLong2.mul.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 9223372036854775807^: r = vA / qB: out = out & "var.vEmp.div.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 9223372036854775807^: r = qB / vA: out = out & "var.LongLong2.div.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 9223372036854775807^: r = vA \ qB: out = out & "var.vEmp.idiv.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 9223372036854775807^: r = qB \ vA: out = out & "var.LongLong2.idiv.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 9223372036854775807^: r = vA Mod qB: out = out & "var.vEmp.mod.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 9223372036854775807^: r = qB Mod vA: out = out & "var.LongLong2.mod.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 9223372036854775807^: r = vA ^ qB: out = out & "var.vEmp.pow.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: qB = 9223372036854775807^: r = qB ^ vA: out = out & "var.LongLong2.pow.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 0.5: r = vA + sB: out = out & "var.vEmp.add.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 0.5: r = sB + vA: out = out & "var.Single0.add.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 0.5: r = vA - sB: out = out & "var.vEmp.sub.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 0.5: r = sB - vA: out = out & "var.Single0.sub.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 0.5: r = vA * sB: out = out & "var.vEmp.mul.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 0.5: r = sB * vA: out = out & "var.Single0.mul.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 0.5: r = vA / sB: out = out & "var.vEmp.div.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 0.5: r = sB / vA: out = out & "var.Single0.div.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 0.5: r = vA \ sB: out = out & "var.vEmp.idiv.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 0.5: r = sB \ vA: out = out & "var.Single0.idiv.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 0.5: r = vA Mod sB: out = out & "var.vEmp.mod.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 0.5: r = sB Mod vA: out = out & "var.Single0.mod.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 0.5: r = vA ^ sB: out = out & "var.vEmp.pow.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 0.5: r = sB ^ vA: out = out & "var.Single0.pow.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 16777216: r = vA + sB: out = out & "var.vEmp.add.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 16777216: r = sB + vA: out = out & "var.Single1.add.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 16777216: r = vA - sB: out = out & "var.vEmp.sub.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 16777216: r = sB - vA: out = out & "var.Single1.sub.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 16777216: r = vA * sB: out = out & "var.vEmp.mul.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 16777216: r = sB * vA: out = out & "var.Single1.mul.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 16777216: r = vA / sB: out = out & "var.vEmp.div.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 16777216: r = sB / vA: out = out & "var.Single1.div.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 16777216: r = vA \ sB: out = out & "var.vEmp.idiv.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 16777216: r = sB \ vA: out = out & "var.Single1.idiv.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 16777216: r = vA Mod sB: out = out & "var.vEmp.mod.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 16777216: r = sB Mod vA: out = out & "var.Single1.mod.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 16777216: r = vA ^ sB: out = out & "var.vEmp.pow.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 16777216: r = sB ^ vA: out = out & "var.Single1.pow.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 3.0E+38: r = vA + sB: out = out & "var.vEmp.add.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 3.0E+38: r = sB + vA: out = out & "var.Single2.add.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 3.0E+38: r = vA - sB: out = out & "var.vEmp.sub.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 3.0E+38: r = sB - vA: out = out & "var.Single2.sub.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 3.0E+38: r = vA * sB: out = out & "var.vEmp.mul.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 3.0E+38: r = sB * vA: out = out & "var.Single2.mul.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 3.0E+38: r = vA / sB: out = out & "var.vEmp.div.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 3.0E+38: r = sB / vA: out = out & "var.Single2.div.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 3.0E+38: r = vA \ sB: out = out & "var.vEmp.idiv.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 3.0E+38: r = sB \ vA: out = out & "var.Single2.idiv.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 3.0E+38: r = vA Mod sB: out = out & "var.vEmp.mod.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 3.0E+38: r = sB Mod vA: out = out & "var.Single2.mod.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 3.0E+38: r = vA ^ sB: out = out & "var.vEmp.pow.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: sB = 3.0E+38: r = sB ^ vA: out = out & "var.Single2.pow.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 0.5: r = vA + dB: out = out & "var.vEmp.add.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 0.5: r = dB + vA: out = out & "var.Double0.add.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 0.5: r = vA - dB: out = out & "var.vEmp.sub.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 0.5: r = dB - vA: out = out & "var.Double0.sub.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 0.5: r = vA * dB: out = out & "var.vEmp.mul.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 0.5: r = dB * vA: out = out & "var.Double0.mul.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 0.5: r = vA / dB: out = out & "var.vEmp.div.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 0.5: r = dB / vA: out = out & "var.Double0.div.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 0.5: r = vA \ dB: out = out & "var.vEmp.idiv.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 0.5: r = dB \ vA: out = out & "var.Double0.idiv.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 0.5: r = vA Mod dB: out = out & "var.vEmp.mod.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 0.5: r = dB Mod vA: out = out & "var.Double0.mod.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 0.5: r = vA ^ dB: out = out & "var.vEmp.pow.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 0.5: r = dB ^ vA: out = out & "var.Double0.pow.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 2.5: r = vA + dB: out = out & "var.vEmp.add.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 2.5: r = dB + vA: out = out & "var.Double1.add.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 2.5: r = vA - dB: out = out & "var.vEmp.sub.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 2.5: r = dB - vA: out = out & "var.Double1.sub.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 2.5: r = vA * dB: out = out & "var.vEmp.mul.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 2.5: r = dB * vA: out = out & "var.Double1.mul.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 2.5: r = vA / dB: out = out & "var.vEmp.div.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 2.5: r = dB / vA: out = out & "var.Double1.div.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 2.5: r = vA \ dB: out = out & "var.vEmp.idiv.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 2.5: r = dB \ vA: out = out & "var.Double1.idiv.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 2.5: r = vA Mod dB: out = out & "var.vEmp.mod.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 2.5: r = dB Mod vA: out = out & "var.Double1.mod.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 2.5: r = vA ^ dB: out = out & "var.vEmp.pow.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 2.5: r = dB ^ vA: out = out & "var.Double1.pow.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 1.0E+308: r = vA + dB: out = out & "var.vEmp.add.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 1.0E+308: r = dB + vA: out = out & "var.Double2.add.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 1.0E+308: r = vA - dB: out = out & "var.vEmp.sub.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 1.0E+308: r = dB - vA: out = out & "var.Double2.sub.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 1.0E+308: r = vA * dB: out = out & "var.vEmp.mul.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 1.0E+308: r = dB * vA: out = out & "var.Double2.mul.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 1.0E+308: r = vA / dB: out = out & "var.vEmp.div.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 1.0E+308: r = dB / vA: out = out & "var.Double2.div.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 1.0E+308: r = vA \ dB: out = out & "var.vEmp.idiv.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 1.0E+308: r = dB \ vA: out = out & "var.Double2.idiv.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 1.0E+308: r = vA Mod dB: out = out & "var.vEmp.mod.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 1.0E+308: r = dB Mod vA: out = out & "var.Double2.mod.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 1.0E+308: r = vA ^ dB: out = out & "var.vEmp.pow.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: dB = 1.0E+308: r = dB ^ vA: out = out & "var.Double2.pow.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = -0.5@: r = vA + cB: out = out & "var.vEmp.add.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = -0.5@: r = cB + vA: out = out & "var.Currency0.add.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = -0.5@: r = vA - cB: out = out & "var.vEmp.sub.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = -0.5@: r = cB - vA: out = out & "var.Currency0.sub.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = -0.5@: r = vA * cB: out = out & "var.vEmp.mul.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = -0.5@: r = cB * vA: out = out & "var.Currency0.mul.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = -0.5@: r = vA / cB: out = out & "var.vEmp.div.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = -0.5@: r = cB / vA: out = out & "var.Currency0.div.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = -0.5@: r = vA \ cB: out = out & "var.vEmp.idiv.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = -0.5@: r = cB \ vA: out = out & "var.Currency0.idiv.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = -0.5@: r = vA Mod cB: out = out & "var.vEmp.mod.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = -0.5@: r = cB Mod vA: out = out & "var.Currency0.mod.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = -0.5@: r = vA ^ cB: out = out & "var.vEmp.pow.Currency0=" & Rz(Err.Number, r) & vbLf
  P047 = out
End Function
Function P048() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: vA = Empty: cB = -0.5@: r = cB ^ vA: out = out & "var.Currency0.pow.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 1.5@: r = vA + cB: out = out & "var.vEmp.add.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 1.5@: r = cB + vA: out = out & "var.Currency1.add.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 1.5@: r = vA - cB: out = out & "var.vEmp.sub.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 1.5@: r = cB - vA: out = out & "var.Currency1.sub.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 1.5@: r = vA * cB: out = out & "var.vEmp.mul.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 1.5@: r = cB * vA: out = out & "var.Currency1.mul.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 1.5@: r = vA / cB: out = out & "var.vEmp.div.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 1.5@: r = cB / vA: out = out & "var.Currency1.div.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 1.5@: r = vA \ cB: out = out & "var.vEmp.idiv.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 1.5@: r = cB \ vA: out = out & "var.Currency1.idiv.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 1.5@: r = vA Mod cB: out = out & "var.vEmp.mod.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 1.5@: r = cB Mod vA: out = out & "var.Currency1.mod.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 1.5@: r = vA ^ cB: out = out & "var.vEmp.pow.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 1.5@: r = cB ^ vA: out = out & "var.Currency1.pow.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 922337203685477.5807@: r = vA + cB: out = out & "var.vEmp.add.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 922337203685477.5807@: r = cB + vA: out = out & "var.Currency2.add.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 922337203685477.5807@: r = vA - cB: out = out & "var.vEmp.sub.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 922337203685477.5807@: r = cB - vA: out = out & "var.Currency2.sub.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 922337203685477.5807@: r = vA * cB: out = out & "var.vEmp.mul.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 922337203685477.5807@: r = cB * vA: out = out & "var.Currency2.mul.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 922337203685477.5807@: r = vA / cB: out = out & "var.vEmp.div.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 922337203685477.5807@: r = cB / vA: out = out & "var.Currency2.div.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 922337203685477.5807@: r = vA \ cB: out = out & "var.vEmp.idiv.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 922337203685477.5807@: r = cB \ vA: out = out & "var.Currency2.idiv.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 922337203685477.5807@: r = vA Mod cB: out = out & "var.vEmp.mod.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 922337203685477.5807@: r = cB Mod vA: out = out & "var.Currency2.mod.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 922337203685477.5807@: r = vA ^ cB: out = out & "var.vEmp.pow.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: cB = 922337203685477.5807@: r = cB ^ vA: out = out & "var.Currency2.pow.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: r = -bA: out = out & "neg.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: r = Not bA: out = out & "not.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: r = Abs(bA): out = out & "abs.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: r = -bA: out = out & "neg.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: r = Not bA: out = out & "not.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: r = Abs(bA): out = out & "abs.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: r = -bA: out = out & "neg.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: r = Not bA: out = out & "not.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: r = Abs(bA): out = out & "abs.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: r = -iA: out = out & "neg.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: r = Not iA: out = out & "not.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: r = Abs(iA): out = out & "abs.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: r = -iA: out = out & "neg.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: r = Not iA: out = out & "not.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: r = Abs(iA): out = out & "abs.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: r = -iA: out = out & "neg.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: r = Not iA: out = out & "not.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: r = Abs(iA): out = out & "abs.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: r = -lA: out = out & "neg.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: r = Not lA: out = out & "not.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: r = Abs(lA): out = out & "abs.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: r = -lA: out = out & "neg.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: r = Not lA: out = out & "not.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: r = Abs(lA): out = out & "abs.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: r = -lA: out = out & "neg.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: r = Not lA: out = out & "not.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: r = Abs(lA): out = out & "abs.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: r = -qA: out = out & "neg.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: r = Not qA: out = out & "not.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: r = Abs(qA): out = out & "abs.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: r = -qA: out = out & "neg.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: r = Not qA: out = out & "not.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: r = Abs(qA): out = out & "abs.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: r = -qA: out = out & "neg.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: r = Not qA: out = out & "not.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: r = Abs(qA): out = out & "abs.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: r = -sA: out = out & "neg.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: r = Not sA: out = out & "not.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: r = Abs(sA): out = out & "abs.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: r = -sA: out = out & "neg.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: r = Not sA: out = out & "not.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: r = Abs(sA): out = out & "abs.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: r = -sA: out = out & "neg.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: r = Not sA: out = out & "not.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: r = Abs(sA): out = out & "abs.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: r = -dA: out = out & "neg.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: r = Not dA: out = out & "not.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: r = Abs(dA): out = out & "abs.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: r = -dA: out = out & "neg.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: r = Not dA: out = out & "not.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: r = Abs(dA): out = out & "abs.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: r = -dA: out = out & "neg.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: r = Not dA: out = out & "not.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: r = Abs(dA): out = out & "abs.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: r = -cA: out = out & "neg.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: r = Not cA: out = out & "not.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: r = Abs(cA): out = out & "abs.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: r = -cA: out = out & "neg.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: r = Not cA: out = out & "not.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: r = Abs(cA): out = out & "abs.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: r = -cA: out = out & "neg.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: r = Not cA: out = out & "not.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: r = Abs(cA): out = out & "abs.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: r = -vA: out = out & "neg.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: r = Not vA: out = out & "not.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: r = Abs(vA): out = out & "abs.Variant0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: r = -vA: out = out & "neg.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: r = Not vA: out = out & "not.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 7: r = Abs(vA): out = out & "abs.Variant1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: r = -vA: out = out & "neg.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: r = Not vA: out = out & "not.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: r = Abs(vA): out = out & "abs.Variant2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: bR = bA: out = out & "store.Byte.from.Byte0=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: bA = 0: bR = bA + 1: out = out & "store.Byte.add1.Byte0=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: bA = 0: bR = bA * 2: out = out & "store.Byte.mul2.Byte0=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: bA = 200: bR = bA: out = out & "store.Byte.from.Byte1=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: bA = 200: bR = bA + 1: out = out & "store.Byte.add1.Byte1=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: bA = 200: bR = bA * 2: out = out & "store.Byte.mul2.Byte1=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: bA = 255: bR = bA: out = out & "store.Byte.from.Byte2=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: bA = 255: bR = bA + 1: out = out & "store.Byte.add1.Byte2=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: bA = 255: bR = bA * 2: out = out & "store.Byte.mul2.Byte2=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: iA = -32768: bR = iA: out = out & "store.Byte.from.Integer0=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: iA = -32768: bR = iA + 1: out = out & "store.Byte.add1.Integer0=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: iA = -32768: bR = iA * 2: out = out & "store.Byte.mul2.Integer0=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: iA = 1: bR = iA: out = out & "store.Byte.from.Integer1=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: iA = 1: bR = iA + 1: out = out & "store.Byte.add1.Integer1=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: iA = 1: bR = iA * 2: out = out & "store.Byte.mul2.Integer1=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: iA = 32767: bR = iA: out = out & "store.Byte.from.Integer2=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: iA = 32767: bR = iA + 1: out = out & "store.Byte.add1.Integer2=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: iA = 32767: bR = iA * 2: out = out & "store.Byte.mul2.Integer2=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: lA = -2147483648: bR = lA: out = out & "store.Byte.from.Long0=" & Rz(Err.Number, bR) & vbLf
  P048 = out
End Function
Function P049() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: lA = -2147483648: bR = lA + 1: out = out & "store.Byte.add1.Long0=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: lA = -2147483648: bR = lA * 2: out = out & "store.Byte.mul2.Long0=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: lA = 65536: bR = lA: out = out & "store.Byte.from.Long1=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: lA = 65536: bR = lA + 1: out = out & "store.Byte.add1.Long1=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: lA = 65536: bR = lA * 2: out = out & "store.Byte.mul2.Long1=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: lA = 2147483647: bR = lA: out = out & "store.Byte.from.Long2=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: lA = 2147483647: bR = lA + 1: out = out & "store.Byte.add1.Long2=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: lA = 2147483647: bR = lA * 2: out = out & "store.Byte.mul2.Long2=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: sA = 0.5: bR = sA: out = out & "store.Byte.from.Single0=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: sA = 0.5: bR = sA + 1: out = out & "store.Byte.add1.Single0=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: sA = 0.5: bR = sA * 2: out = out & "store.Byte.mul2.Single0=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: sA = 16777216: bR = sA: out = out & "store.Byte.from.Single1=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: sA = 16777216: bR = sA + 1: out = out & "store.Byte.add1.Single1=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: sA = 16777216: bR = sA * 2: out = out & "store.Byte.mul2.Single1=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: sA = 3.0E+38: bR = sA: out = out & "store.Byte.from.Single2=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: sA = 3.0E+38: bR = sA + 1: out = out & "store.Byte.add1.Single2=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: sA = 3.0E+38: bR = sA * 2: out = out & "store.Byte.mul2.Single2=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: dA = 0.5: bR = dA: out = out & "store.Byte.from.Double0=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: dA = 0.5: bR = dA + 1: out = out & "store.Byte.add1.Double0=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: dA = 0.5: bR = dA * 2: out = out & "store.Byte.mul2.Double0=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: dA = 2.5: bR = dA: out = out & "store.Byte.from.Double1=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: dA = 2.5: bR = dA + 1: out = out & "store.Byte.add1.Double1=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: dA = 2.5: bR = dA * 2: out = out & "store.Byte.mul2.Double1=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: dA = 1.0E+308: bR = dA: out = out & "store.Byte.from.Double2=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: dA = 1.0E+308: bR = dA + 1: out = out & "store.Byte.add1.Double2=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: dA = 1.0E+308: bR = dA * 2: out = out & "store.Byte.mul2.Double2=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: cA = -0.5@: bR = cA: out = out & "store.Byte.from.Currency0=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: cA = -0.5@: bR = cA + 1: out = out & "store.Byte.add1.Currency0=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: cA = -0.5@: bR = cA * 2: out = out & "store.Byte.mul2.Currency0=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: cA = 1.5@: bR = cA: out = out & "store.Byte.from.Currency1=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: cA = 1.5@: bR = cA + 1: out = out & "store.Byte.add1.Currency1=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: cA = 1.5@: bR = cA * 2: out = out & "store.Byte.mul2.Currency1=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bR = cA: out = out & "store.Byte.from.Currency2=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bR = cA + 1: out = out & "store.Byte.add1.Currency2=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bR = cA * 2: out = out & "store.Byte.mul2.Currency2=" & Rz(Err.Number, bR) & vbLf
  Err.Clear: bA = 0: iR = bA: out = out & "store.Integer.from.Byte0=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: bA = 0: iR = bA + 1: out = out & "store.Integer.add1.Byte0=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: bA = 0: iR = bA * 2: out = out & "store.Integer.mul2.Byte0=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: bA = 200: iR = bA: out = out & "store.Integer.from.Byte1=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: bA = 200: iR = bA + 1: out = out & "store.Integer.add1.Byte1=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: bA = 200: iR = bA * 2: out = out & "store.Integer.mul2.Byte1=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: bA = 255: iR = bA: out = out & "store.Integer.from.Byte2=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: bA = 255: iR = bA + 1: out = out & "store.Integer.add1.Byte2=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: bA = 255: iR = bA * 2: out = out & "store.Integer.mul2.Byte2=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: iA = -32768: iR = iA: out = out & "store.Integer.from.Integer0=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: iA = -32768: iR = iA + 1: out = out & "store.Integer.add1.Integer0=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: iA = -32768: iR = iA * 2: out = out & "store.Integer.mul2.Integer0=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: iA = 1: iR = iA: out = out & "store.Integer.from.Integer1=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: iA = 1: iR = iA + 1: out = out & "store.Integer.add1.Integer1=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: iA = 1: iR = iA * 2: out = out & "store.Integer.mul2.Integer1=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: iA = 32767: iR = iA: out = out & "store.Integer.from.Integer2=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: iA = 32767: iR = iA + 1: out = out & "store.Integer.add1.Integer2=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: iA = 32767: iR = iA * 2: out = out & "store.Integer.mul2.Integer2=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: lA = -2147483648: iR = lA: out = out & "store.Integer.from.Long0=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: lA = -2147483648: iR = lA + 1: out = out & "store.Integer.add1.Long0=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: lA = -2147483648: iR = lA * 2: out = out & "store.Integer.mul2.Long0=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: lA = 65536: iR = lA: out = out & "store.Integer.from.Long1=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: lA = 65536: iR = lA + 1: out = out & "store.Integer.add1.Long1=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: lA = 65536: iR = lA * 2: out = out & "store.Integer.mul2.Long1=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: lA = 2147483647: iR = lA: out = out & "store.Integer.from.Long2=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: lA = 2147483647: iR = lA + 1: out = out & "store.Integer.add1.Long2=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: lA = 2147483647: iR = lA * 2: out = out & "store.Integer.mul2.Long2=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: sA = 0.5: iR = sA: out = out & "store.Integer.from.Single0=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: sA = 0.5: iR = sA + 1: out = out & "store.Integer.add1.Single0=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: sA = 0.5: iR = sA * 2: out = out & "store.Integer.mul2.Single0=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: sA = 16777216: iR = sA: out = out & "store.Integer.from.Single1=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: sA = 16777216: iR = sA + 1: out = out & "store.Integer.add1.Single1=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: sA = 16777216: iR = sA * 2: out = out & "store.Integer.mul2.Single1=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: sA = 3.0E+38: iR = sA: out = out & "store.Integer.from.Single2=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: sA = 3.0E+38: iR = sA + 1: out = out & "store.Integer.add1.Single2=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: sA = 3.0E+38: iR = sA * 2: out = out & "store.Integer.mul2.Single2=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: dA = 0.5: iR = dA: out = out & "store.Integer.from.Double0=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: dA = 0.5: iR = dA + 1: out = out & "store.Integer.add1.Double0=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: dA = 0.5: iR = dA * 2: out = out & "store.Integer.mul2.Double0=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: dA = 2.5: iR = dA: out = out & "store.Integer.from.Double1=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: dA = 2.5: iR = dA + 1: out = out & "store.Integer.add1.Double1=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: dA = 2.5: iR = dA * 2: out = out & "store.Integer.mul2.Double1=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: dA = 1.0E+308: iR = dA: out = out & "store.Integer.from.Double2=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: dA = 1.0E+308: iR = dA + 1: out = out & "store.Integer.add1.Double2=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: dA = 1.0E+308: iR = dA * 2: out = out & "store.Integer.mul2.Double2=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: cA = -0.5@: iR = cA: out = out & "store.Integer.from.Currency0=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: cA = -0.5@: iR = cA + 1: out = out & "store.Integer.add1.Currency0=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: cA = -0.5@: iR = cA * 2: out = out & "store.Integer.mul2.Currency0=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: cA = 1.5@: iR = cA: out = out & "store.Integer.from.Currency1=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: cA = 1.5@: iR = cA + 1: out = out & "store.Integer.add1.Currency1=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: cA = 1.5@: iR = cA * 2: out = out & "store.Integer.mul2.Currency1=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iR = cA: out = out & "store.Integer.from.Currency2=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iR = cA + 1: out = out & "store.Integer.add1.Currency2=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iR = cA * 2: out = out & "store.Integer.mul2.Currency2=" & Rz(Err.Number, iR) & vbLf
  Err.Clear: bA = 0: lR = bA: out = out & "store.Long.from.Byte0=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: bA = 0: lR = bA + 1: out = out & "store.Long.add1.Byte0=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: bA = 0: lR = bA * 2: out = out & "store.Long.mul2.Byte0=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: bA = 200: lR = bA: out = out & "store.Long.from.Byte1=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: bA = 200: lR = bA + 1: out = out & "store.Long.add1.Byte1=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: bA = 200: lR = bA * 2: out = out & "store.Long.mul2.Byte1=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: bA = 255: lR = bA: out = out & "store.Long.from.Byte2=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: bA = 255: lR = bA + 1: out = out & "store.Long.add1.Byte2=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: bA = 255: lR = bA * 2: out = out & "store.Long.mul2.Byte2=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: iA = -32768: lR = iA: out = out & "store.Long.from.Integer0=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: iA = -32768: lR = iA + 1: out = out & "store.Long.add1.Integer0=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: iA = -32768: lR = iA * 2: out = out & "store.Long.mul2.Integer0=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: iA = 1: lR = iA: out = out & "store.Long.from.Integer1=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: iA = 1: lR = iA + 1: out = out & "store.Long.add1.Integer1=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: iA = 1: lR = iA * 2: out = out & "store.Long.mul2.Integer1=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: iA = 32767: lR = iA: out = out & "store.Long.from.Integer2=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: iA = 32767: lR = iA + 1: out = out & "store.Long.add1.Integer2=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: iA = 32767: lR = iA * 2: out = out & "store.Long.mul2.Integer2=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: lA = -2147483648: lR = lA: out = out & "store.Long.from.Long0=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: lA = -2147483648: lR = lA + 1: out = out & "store.Long.add1.Long0=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: lA = -2147483648: lR = lA * 2: out = out & "store.Long.mul2.Long0=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: lA = 65536: lR = lA: out = out & "store.Long.from.Long1=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: lA = 65536: lR = lA + 1: out = out & "store.Long.add1.Long1=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: lA = 65536: lR = lA * 2: out = out & "store.Long.mul2.Long1=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: lA = 2147483647: lR = lA: out = out & "store.Long.from.Long2=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: lA = 2147483647: lR = lA + 1: out = out & "store.Long.add1.Long2=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: lA = 2147483647: lR = lA * 2: out = out & "store.Long.mul2.Long2=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: sA = 0.5: lR = sA: out = out & "store.Long.from.Single0=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: sA = 0.5: lR = sA + 1: out = out & "store.Long.add1.Single0=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: sA = 0.5: lR = sA * 2: out = out & "store.Long.mul2.Single0=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: sA = 16777216: lR = sA: out = out & "store.Long.from.Single1=" & Rz(Err.Number, lR) & vbLf
  P049 = out
End Function
Function P050() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: sA = 16777216: lR = sA + 1: out = out & "store.Long.add1.Single1=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: sA = 16777216: lR = sA * 2: out = out & "store.Long.mul2.Single1=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: sA = 3.0E+38: lR = sA: out = out & "store.Long.from.Single2=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: sA = 3.0E+38: lR = sA + 1: out = out & "store.Long.add1.Single2=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: sA = 3.0E+38: lR = sA * 2: out = out & "store.Long.mul2.Single2=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: dA = 0.5: lR = dA: out = out & "store.Long.from.Double0=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: dA = 0.5: lR = dA + 1: out = out & "store.Long.add1.Double0=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: dA = 0.5: lR = dA * 2: out = out & "store.Long.mul2.Double0=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: dA = 2.5: lR = dA: out = out & "store.Long.from.Double1=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: dA = 2.5: lR = dA + 1: out = out & "store.Long.add1.Double1=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: dA = 2.5: lR = dA * 2: out = out & "store.Long.mul2.Double1=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: dA = 1.0E+308: lR = dA: out = out & "store.Long.from.Double2=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: dA = 1.0E+308: lR = dA + 1: out = out & "store.Long.add1.Double2=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: dA = 1.0E+308: lR = dA * 2: out = out & "store.Long.mul2.Double2=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: cA = -0.5@: lR = cA: out = out & "store.Long.from.Currency0=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: cA = -0.5@: lR = cA + 1: out = out & "store.Long.add1.Currency0=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: cA = -0.5@: lR = cA * 2: out = out & "store.Long.mul2.Currency0=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: cA = 1.5@: lR = cA: out = out & "store.Long.from.Currency1=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: cA = 1.5@: lR = cA + 1: out = out & "store.Long.add1.Currency1=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: cA = 1.5@: lR = cA * 2: out = out & "store.Long.mul2.Currency1=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lR = cA: out = out & "store.Long.from.Currency2=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lR = cA + 1: out = out & "store.Long.add1.Currency2=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lR = cA * 2: out = out & "store.Long.mul2.Currency2=" & Rz(Err.Number, lR) & vbLf
  Err.Clear: bA = 0: qR = bA: out = out & "store.LongLong.from.Byte0=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: bA = 0: qR = bA + 1: out = out & "store.LongLong.add1.Byte0=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: bA = 0: qR = bA * 2: out = out & "store.LongLong.mul2.Byte0=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: bA = 200: qR = bA: out = out & "store.LongLong.from.Byte1=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: bA = 200: qR = bA + 1: out = out & "store.LongLong.add1.Byte1=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: bA = 200: qR = bA * 2: out = out & "store.LongLong.mul2.Byte1=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: bA = 255: qR = bA: out = out & "store.LongLong.from.Byte2=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: bA = 255: qR = bA + 1: out = out & "store.LongLong.add1.Byte2=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: bA = 255: qR = bA * 2: out = out & "store.LongLong.mul2.Byte2=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: iA = -32768: qR = iA: out = out & "store.LongLong.from.Integer0=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: iA = -32768: qR = iA + 1: out = out & "store.LongLong.add1.Integer0=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: iA = -32768: qR = iA * 2: out = out & "store.LongLong.mul2.Integer0=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: iA = 1: qR = iA: out = out & "store.LongLong.from.Integer1=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: iA = 1: qR = iA + 1: out = out & "store.LongLong.add1.Integer1=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: iA = 1: qR = iA * 2: out = out & "store.LongLong.mul2.Integer1=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: iA = 32767: qR = iA: out = out & "store.LongLong.from.Integer2=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: iA = 32767: qR = iA + 1: out = out & "store.LongLong.add1.Integer2=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: iA = 32767: qR = iA * 2: out = out & "store.LongLong.mul2.Integer2=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: lA = -2147483648: qR = lA: out = out & "store.LongLong.from.Long0=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: lA = -2147483648: qR = lA + 1: out = out & "store.LongLong.add1.Long0=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: lA = -2147483648: qR = lA * 2: out = out & "store.LongLong.mul2.Long0=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: lA = 65536: qR = lA: out = out & "store.LongLong.from.Long1=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: lA = 65536: qR = lA + 1: out = out & "store.LongLong.add1.Long1=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: lA = 65536: qR = lA * 2: out = out & "store.LongLong.mul2.Long1=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: lA = 2147483647: qR = lA: out = out & "store.LongLong.from.Long2=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: lA = 2147483647: qR = lA + 1: out = out & "store.LongLong.add1.Long2=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: lA = 2147483647: qR = lA * 2: out = out & "store.LongLong.mul2.Long2=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: qA = -9223372036854775807^: qR = qA: out = out & "store.LongLong.from.LongLong0=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: qA = -9223372036854775807^: qR = qA + 1: out = out & "store.LongLong.add1.LongLong0=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: qA = -9223372036854775807^: qR = qA * 2: out = out & "store.LongLong.mul2.LongLong0=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: qA = 3: qR = qA: out = out & "store.LongLong.from.LongLong1=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: qA = 3: qR = qA + 1: out = out & "store.LongLong.add1.LongLong1=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: qA = 3: qR = qA * 2: out = out & "store.LongLong.mul2.LongLong1=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: qA = 9223372036854775807^: qR = qA: out = out & "store.LongLong.from.LongLong2=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: qA = 9223372036854775807^: qR = qA + 1: out = out & "store.LongLong.add1.LongLong2=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: qA = 9223372036854775807^: qR = qA * 2: out = out & "store.LongLong.mul2.LongLong2=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: sA = 0.5: qR = sA: out = out & "store.LongLong.from.Single0=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: sA = 0.5: qR = sA + 1: out = out & "store.LongLong.add1.Single0=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: sA = 0.5: qR = sA * 2: out = out & "store.LongLong.mul2.Single0=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: sA = 16777216: qR = sA: out = out & "store.LongLong.from.Single1=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: sA = 16777216: qR = sA + 1: out = out & "store.LongLong.add1.Single1=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: sA = 16777216: qR = sA * 2: out = out & "store.LongLong.mul2.Single1=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: sA = 3.0E+38: qR = sA: out = out & "store.LongLong.from.Single2=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: sA = 3.0E+38: qR = sA + 1: out = out & "store.LongLong.add1.Single2=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: sA = 3.0E+38: qR = sA * 2: out = out & "store.LongLong.mul2.Single2=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: dA = 0.5: qR = dA: out = out & "store.LongLong.from.Double0=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: dA = 0.5: qR = dA + 1: out = out & "store.LongLong.add1.Double0=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: dA = 0.5: qR = dA * 2: out = out & "store.LongLong.mul2.Double0=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: dA = 2.5: qR = dA: out = out & "store.LongLong.from.Double1=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: dA = 2.5: qR = dA + 1: out = out & "store.LongLong.add1.Double1=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: dA = 2.5: qR = dA * 2: out = out & "store.LongLong.mul2.Double1=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: dA = 1.0E+308: qR = dA: out = out & "store.LongLong.from.Double2=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: dA = 1.0E+308: qR = dA + 1: out = out & "store.LongLong.add1.Double2=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: dA = 1.0E+308: qR = dA * 2: out = out & "store.LongLong.mul2.Double2=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: cA = -0.5@: qR = cA: out = out & "store.LongLong.from.Currency0=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: cA = -0.5@: qR = cA + 1: out = out & "store.LongLong.add1.Currency0=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: cA = -0.5@: qR = cA * 2: out = out & "store.LongLong.mul2.Currency0=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: cA = 1.5@: qR = cA: out = out & "store.LongLong.from.Currency1=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: cA = 1.5@: qR = cA + 1: out = out & "store.LongLong.add1.Currency1=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: cA = 1.5@: qR = cA * 2: out = out & "store.LongLong.mul2.Currency1=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qR = cA: out = out & "store.LongLong.from.Currency2=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qR = cA + 1: out = out & "store.LongLong.add1.Currency2=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qR = cA * 2: out = out & "store.LongLong.mul2.Currency2=" & Rz(Err.Number, qR) & vbLf
  Err.Clear: bA = 0: sR = bA: out = out & "store.Single.from.Byte0=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: bA = 0: sR = bA + 1: out = out & "store.Single.add1.Byte0=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: bA = 0: sR = bA * 2: out = out & "store.Single.mul2.Byte0=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: bA = 200: sR = bA: out = out & "store.Single.from.Byte1=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: bA = 200: sR = bA + 1: out = out & "store.Single.add1.Byte1=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: bA = 200: sR = bA * 2: out = out & "store.Single.mul2.Byte1=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: bA = 255: sR = bA: out = out & "store.Single.from.Byte2=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: bA = 255: sR = bA + 1: out = out & "store.Single.add1.Byte2=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: bA = 255: sR = bA * 2: out = out & "store.Single.mul2.Byte2=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: iA = -32768: sR = iA: out = out & "store.Single.from.Integer0=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: iA = -32768: sR = iA + 1: out = out & "store.Single.add1.Integer0=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: iA = -32768: sR = iA * 2: out = out & "store.Single.mul2.Integer0=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: iA = 1: sR = iA: out = out & "store.Single.from.Integer1=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: iA = 1: sR = iA + 1: out = out & "store.Single.add1.Integer1=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: iA = 1: sR = iA * 2: out = out & "store.Single.mul2.Integer1=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: iA = 32767: sR = iA: out = out & "store.Single.from.Integer2=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: iA = 32767: sR = iA + 1: out = out & "store.Single.add1.Integer2=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: iA = 32767: sR = iA * 2: out = out & "store.Single.mul2.Integer2=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: lA = -2147483648: sR = lA: out = out & "store.Single.from.Long0=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: lA = -2147483648: sR = lA + 1: out = out & "store.Single.add1.Long0=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: lA = -2147483648: sR = lA * 2: out = out & "store.Single.mul2.Long0=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: lA = 65536: sR = lA: out = out & "store.Single.from.Long1=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: lA = 65536: sR = lA + 1: out = out & "store.Single.add1.Long1=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: lA = 65536: sR = lA * 2: out = out & "store.Single.mul2.Long1=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: lA = 2147483647: sR = lA: out = out & "store.Single.from.Long2=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: lA = 2147483647: sR = lA + 1: out = out & "store.Single.add1.Long2=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: lA = 2147483647: sR = lA * 2: out = out & "store.Single.mul2.Long2=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: qA = -9223372036854775807^: sR = qA: out = out & "store.Single.from.LongLong0=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: qA = -9223372036854775807^: sR = qA + 1: out = out & "store.Single.add1.LongLong0=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: qA = -9223372036854775807^: sR = qA * 2: out = out & "store.Single.mul2.LongLong0=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: qA = 3: sR = qA: out = out & "store.Single.from.LongLong1=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: qA = 3: sR = qA + 1: out = out & "store.Single.add1.LongLong1=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: qA = 3: sR = qA * 2: out = out & "store.Single.mul2.LongLong1=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: qA = 9223372036854775807^: sR = qA: out = out & "store.Single.from.LongLong2=" & Rz(Err.Number, sR) & vbLf
  P050 = out
End Function
Function P051() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: qA = 9223372036854775807^: sR = qA + 1: out = out & "store.Single.add1.LongLong2=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: qA = 9223372036854775807^: sR = qA * 2: out = out & "store.Single.mul2.LongLong2=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: sA = 0.5: sR = sA: out = out & "store.Single.from.Single0=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: sA = 0.5: sR = sA + 1: out = out & "store.Single.add1.Single0=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: sA = 0.5: sR = sA * 2: out = out & "store.Single.mul2.Single0=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: sA = 16777216: sR = sA: out = out & "store.Single.from.Single1=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: sA = 16777216: sR = sA + 1: out = out & "store.Single.add1.Single1=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: sA = 16777216: sR = sA * 2: out = out & "store.Single.mul2.Single1=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: sA = 3.0E+38: sR = sA: out = out & "store.Single.from.Single2=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: sA = 3.0E+38: sR = sA + 1: out = out & "store.Single.add1.Single2=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: sA = 3.0E+38: sR = sA * 2: out = out & "store.Single.mul2.Single2=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: dA = 0.5: sR = dA: out = out & "store.Single.from.Double0=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: dA = 0.5: sR = dA + 1: out = out & "store.Single.add1.Double0=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: dA = 0.5: sR = dA * 2: out = out & "store.Single.mul2.Double0=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: dA = 2.5: sR = dA: out = out & "store.Single.from.Double1=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: dA = 2.5: sR = dA + 1: out = out & "store.Single.add1.Double1=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: dA = 2.5: sR = dA * 2: out = out & "store.Single.mul2.Double1=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: dA = 1.0E+308: sR = dA: out = out & "store.Single.from.Double2=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: dA = 1.0E+308: sR = dA + 1: out = out & "store.Single.add1.Double2=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: dA = 1.0E+308: sR = dA * 2: out = out & "store.Single.mul2.Double2=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: cA = -0.5@: sR = cA: out = out & "store.Single.from.Currency0=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: cA = -0.5@: sR = cA + 1: out = out & "store.Single.add1.Currency0=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: cA = -0.5@: sR = cA * 2: out = out & "store.Single.mul2.Currency0=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: cA = 1.5@: sR = cA: out = out & "store.Single.from.Currency1=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: cA = 1.5@: sR = cA + 1: out = out & "store.Single.add1.Currency1=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: cA = 1.5@: sR = cA * 2: out = out & "store.Single.mul2.Currency1=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sR = cA: out = out & "store.Single.from.Currency2=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sR = cA + 1: out = out & "store.Single.add1.Currency2=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sR = cA * 2: out = out & "store.Single.mul2.Currency2=" & Rz(Err.Number, sR) & vbLf
  Err.Clear: bA = 0: dR = bA: out = out & "store.Double.from.Byte0=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: bA = 0: dR = bA + 1: out = out & "store.Double.add1.Byte0=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: bA = 0: dR = bA * 2: out = out & "store.Double.mul2.Byte0=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: bA = 200: dR = bA: out = out & "store.Double.from.Byte1=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: bA = 200: dR = bA + 1: out = out & "store.Double.add1.Byte1=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: bA = 200: dR = bA * 2: out = out & "store.Double.mul2.Byte1=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: bA = 255: dR = bA: out = out & "store.Double.from.Byte2=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: bA = 255: dR = bA + 1: out = out & "store.Double.add1.Byte2=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: bA = 255: dR = bA * 2: out = out & "store.Double.mul2.Byte2=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: iA = -32768: dR = iA: out = out & "store.Double.from.Integer0=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: iA = -32768: dR = iA + 1: out = out & "store.Double.add1.Integer0=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: iA = -32768: dR = iA * 2: out = out & "store.Double.mul2.Integer0=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: iA = 1: dR = iA: out = out & "store.Double.from.Integer1=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: iA = 1: dR = iA + 1: out = out & "store.Double.add1.Integer1=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: iA = 1: dR = iA * 2: out = out & "store.Double.mul2.Integer1=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: iA = 32767: dR = iA: out = out & "store.Double.from.Integer2=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: iA = 32767: dR = iA + 1: out = out & "store.Double.add1.Integer2=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: iA = 32767: dR = iA * 2: out = out & "store.Double.mul2.Integer2=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: lA = -2147483648: dR = lA: out = out & "store.Double.from.Long0=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: lA = -2147483648: dR = lA + 1: out = out & "store.Double.add1.Long0=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: lA = -2147483648: dR = lA * 2: out = out & "store.Double.mul2.Long0=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: lA = 65536: dR = lA: out = out & "store.Double.from.Long1=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: lA = 65536: dR = lA + 1: out = out & "store.Double.add1.Long1=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: lA = 65536: dR = lA * 2: out = out & "store.Double.mul2.Long1=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: lA = 2147483647: dR = lA: out = out & "store.Double.from.Long2=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: lA = 2147483647: dR = lA + 1: out = out & "store.Double.add1.Long2=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: lA = 2147483647: dR = lA * 2: out = out & "store.Double.mul2.Long2=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: qA = -9223372036854775807^: dR = qA: out = out & "store.Double.from.LongLong0=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: qA = -9223372036854775807^: dR = qA + 1: out = out & "store.Double.add1.LongLong0=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: qA = -9223372036854775807^: dR = qA * 2: out = out & "store.Double.mul2.LongLong0=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: qA = 3: dR = qA: out = out & "store.Double.from.LongLong1=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: qA = 3: dR = qA + 1: out = out & "store.Double.add1.LongLong1=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: qA = 3: dR = qA * 2: out = out & "store.Double.mul2.LongLong1=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: qA = 9223372036854775807^: dR = qA: out = out & "store.Double.from.LongLong2=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: qA = 9223372036854775807^: dR = qA + 1: out = out & "store.Double.add1.LongLong2=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: qA = 9223372036854775807^: dR = qA * 2: out = out & "store.Double.mul2.LongLong2=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: sA = 0.5: dR = sA: out = out & "store.Double.from.Single0=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: sA = 0.5: dR = sA + 1: out = out & "store.Double.add1.Single0=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: sA = 0.5: dR = sA * 2: out = out & "store.Double.mul2.Single0=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: sA = 16777216: dR = sA: out = out & "store.Double.from.Single1=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: sA = 16777216: dR = sA + 1: out = out & "store.Double.add1.Single1=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: sA = 16777216: dR = sA * 2: out = out & "store.Double.mul2.Single1=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: sA = 3.0E+38: dR = sA: out = out & "store.Double.from.Single2=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: sA = 3.0E+38: dR = sA + 1: out = out & "store.Double.add1.Single2=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: sA = 3.0E+38: dR = sA * 2: out = out & "store.Double.mul2.Single2=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: dA = 0.5: dR = dA: out = out & "store.Double.from.Double0=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: dA = 0.5: dR = dA + 1: out = out & "store.Double.add1.Double0=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: dA = 0.5: dR = dA * 2: out = out & "store.Double.mul2.Double0=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: dA = 2.5: dR = dA: out = out & "store.Double.from.Double1=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: dA = 2.5: dR = dA + 1: out = out & "store.Double.add1.Double1=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: dA = 2.5: dR = dA * 2: out = out & "store.Double.mul2.Double1=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: dA = 1.0E+308: dR = dA: out = out & "store.Double.from.Double2=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: dA = 1.0E+308: dR = dA + 1: out = out & "store.Double.add1.Double2=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: dA = 1.0E+308: dR = dA * 2: out = out & "store.Double.mul2.Double2=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: cA = -0.5@: dR = cA: out = out & "store.Double.from.Currency0=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: cA = -0.5@: dR = cA + 1: out = out & "store.Double.add1.Currency0=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: cA = -0.5@: dR = cA * 2: out = out & "store.Double.mul2.Currency0=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: cA = 1.5@: dR = cA: out = out & "store.Double.from.Currency1=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: cA = 1.5@: dR = cA + 1: out = out & "store.Double.add1.Currency1=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: cA = 1.5@: dR = cA * 2: out = out & "store.Double.mul2.Currency1=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dR = cA: out = out & "store.Double.from.Currency2=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dR = cA + 1: out = out & "store.Double.add1.Currency2=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dR = cA * 2: out = out & "store.Double.mul2.Currency2=" & Rz(Err.Number, dR) & vbLf
  Err.Clear: bA = 0: cR = bA: out = out & "store.Currency.from.Byte0=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: bA = 0: cR = bA + 1: out = out & "store.Currency.add1.Byte0=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: bA = 0: cR = bA * 2: out = out & "store.Currency.mul2.Byte0=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: bA = 200: cR = bA: out = out & "store.Currency.from.Byte1=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: bA = 200: cR = bA + 1: out = out & "store.Currency.add1.Byte1=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: bA = 200: cR = bA * 2: out = out & "store.Currency.mul2.Byte1=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: bA = 255: cR = bA: out = out & "store.Currency.from.Byte2=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: bA = 255: cR = bA + 1: out = out & "store.Currency.add1.Byte2=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: bA = 255: cR = bA * 2: out = out & "store.Currency.mul2.Byte2=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: iA = -32768: cR = iA: out = out & "store.Currency.from.Integer0=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: iA = -32768: cR = iA + 1: out = out & "store.Currency.add1.Integer0=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: iA = -32768: cR = iA * 2: out = out & "store.Currency.mul2.Integer0=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: iA = 1: cR = iA: out = out & "store.Currency.from.Integer1=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: iA = 1: cR = iA + 1: out = out & "store.Currency.add1.Integer1=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: iA = 1: cR = iA * 2: out = out & "store.Currency.mul2.Integer1=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: iA = 32767: cR = iA: out = out & "store.Currency.from.Integer2=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: iA = 32767: cR = iA + 1: out = out & "store.Currency.add1.Integer2=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: iA = 32767: cR = iA * 2: out = out & "store.Currency.mul2.Integer2=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: lA = -2147483648: cR = lA: out = out & "store.Currency.from.Long0=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: lA = -2147483648: cR = lA + 1: out = out & "store.Currency.add1.Long0=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: lA = -2147483648: cR = lA * 2: out = out & "store.Currency.mul2.Long0=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: lA = 65536: cR = lA: out = out & "store.Currency.from.Long1=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: lA = 65536: cR = lA + 1: out = out & "store.Currency.add1.Long1=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: lA = 65536: cR = lA * 2: out = out & "store.Currency.mul2.Long1=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: lA = 2147483647: cR = lA: out = out & "store.Currency.from.Long2=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: lA = 2147483647: cR = lA + 1: out = out & "store.Currency.add1.Long2=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: lA = 2147483647: cR = lA * 2: out = out & "store.Currency.mul2.Long2=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: sA = 0.5: cR = sA: out = out & "store.Currency.from.Single0=" & Rz(Err.Number, cR) & vbLf
  P051 = out
End Function
Function P052() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: sA = 0.5: cR = sA + 1: out = out & "store.Currency.add1.Single0=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: sA = 0.5: cR = sA * 2: out = out & "store.Currency.mul2.Single0=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: sA = 16777216: cR = sA: out = out & "store.Currency.from.Single1=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: sA = 16777216: cR = sA + 1: out = out & "store.Currency.add1.Single1=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: sA = 16777216: cR = sA * 2: out = out & "store.Currency.mul2.Single1=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: sA = 3.0E+38: cR = sA: out = out & "store.Currency.from.Single2=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: sA = 3.0E+38: cR = sA + 1: out = out & "store.Currency.add1.Single2=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: sA = 3.0E+38: cR = sA * 2: out = out & "store.Currency.mul2.Single2=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: dA = 0.5: cR = dA: out = out & "store.Currency.from.Double0=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: dA = 0.5: cR = dA + 1: out = out & "store.Currency.add1.Double0=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: dA = 0.5: cR = dA * 2: out = out & "store.Currency.mul2.Double0=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: dA = 2.5: cR = dA: out = out & "store.Currency.from.Double1=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: dA = 2.5: cR = dA + 1: out = out & "store.Currency.add1.Double1=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: dA = 2.5: cR = dA * 2: out = out & "store.Currency.mul2.Double1=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: dA = 1.0E+308: cR = dA: out = out & "store.Currency.from.Double2=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: dA = 1.0E+308: cR = dA + 1: out = out & "store.Currency.add1.Double2=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: dA = 1.0E+308: cR = dA * 2: out = out & "store.Currency.mul2.Double2=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: cA = -0.5@: cR = cA: out = out & "store.Currency.from.Currency0=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: cA = -0.5@: cR = cA + 1: out = out & "store.Currency.add1.Currency0=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: cA = -0.5@: cR = cA * 2: out = out & "store.Currency.mul2.Currency0=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: cA = 1.5@: cR = cA: out = out & "store.Currency.from.Currency1=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: cA = 1.5@: cR = cA + 1: out = out & "store.Currency.add1.Currency1=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: cA = 1.5@: cR = cA * 2: out = out & "store.Currency.mul2.Currency1=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cR = cA: out = out & "store.Currency.from.Currency2=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cR = cA + 1: out = out & "store.Currency.add1.Currency2=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cR = cA * 2: out = out & "store.Currency.mul2.Currency2=" & Rz(Err.Number, cR) & vbLf
  Err.Clear: dA = -0.5: r = CByte(dA): out = out & "conv.CByte.d0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: r = CByte(dA): out = out & "conv.CByte.d1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.5: r = CByte(dA): out = out & "conv.CByte.d2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: r = CByte(dA): out = out & "conv.CByte.d3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 255.5: r = CByte(dA): out = out & "conv.CByte.d4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 256: r = CByte(dA): out = out & "conv.CByte.d5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 32767.5: r = CByte(dA): out = out & "conv.CByte.d6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 32768: r = CByte(dA): out = out & "conv.CByte.d7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2147483647.5: r = CByte(dA): out = out & "conv.CByte.d8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2147483648: r = CByte(dA): out = out & "conv.CByte.d9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 4.5E+18: r = CByte(dA): out = out & "conv.CByte.d10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 9.3E+18: r = CByte(dA): out = out & "conv.CByte.d11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = -1: r = CByte(dA): out = out & "conv.CByte.d12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+38: r = CByte(dA): out = out & "conv.CByte.d13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 3.5E+38: r = CByte(dA): out = out & "conv.CByte.d14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+300: r = CByte(dA): out = out & "conv.CByte.d15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 922337203685477.58: r = CByte(dA): out = out & "conv.CByte.d16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 922337203685478: r = CByte(dA): out = out & "conv.CByte.d17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": r = CByte(vA): out = out & "conv.CByte.s0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "1.5": r = CByte(vA): out = out & "conv.CByte.s1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "2.5": r = CByte(vA): out = out & "conv.CByte.s2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "-0.5": r = CByte(vA): out = out & "conv.CByte.s3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = " 7 ": r = CByte(vA): out = out & "conv.CByte.s4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "1e3": r = CByte(vA): out = out & "conv.CByte.s5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "&H10": r = CByte(vA): out = out & "conv.CByte.s6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "abc": r = CByte(vA): out = out & "conv.CByte.s7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "": r = CByte(vA): out = out & "conv.CByte.s8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "32768": r = CByte(vA): out = out & "conv.CByte.s9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "99999999999": r = CByte(vA): out = out & "conv.CByte.s10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: r = CByte(vA): out = out & "conv.CByte.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: r = CByte(vA): out = out & "conv.CByte.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: r = CByte(vA): out = out & "conv.CByte.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: r = CByte(vA): out = out & "conv.CByte.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": r = CByte(vA): out = out & "conv.CByte.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: r = CByte(vA): out = out & "conv.CByte.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = -0.5: r = CInt(dA): out = out & "conv.CInt.d0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: r = CInt(dA): out = out & "conv.CInt.d1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.5: r = CInt(dA): out = out & "conv.CInt.d2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: r = CInt(dA): out = out & "conv.CInt.d3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 255.5: r = CInt(dA): out = out & "conv.CInt.d4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 256: r = CInt(dA): out = out & "conv.CInt.d5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 32767.5: r = CInt(dA): out = out & "conv.CInt.d6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 32768: r = CInt(dA): out = out & "conv.CInt.d7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2147483647.5: r = CInt(dA): out = out & "conv.CInt.d8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2147483648: r = CInt(dA): out = out & "conv.CInt.d9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 4.5E+18: r = CInt(dA): out = out & "conv.CInt.d10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 9.3E+18: r = CInt(dA): out = out & "conv.CInt.d11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = -1: r = CInt(dA): out = out & "conv.CInt.d12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+38: r = CInt(dA): out = out & "conv.CInt.d13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 3.5E+38: r = CInt(dA): out = out & "conv.CInt.d14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+300: r = CInt(dA): out = out & "conv.CInt.d15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 922337203685477.58: r = CInt(dA): out = out & "conv.CInt.d16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 922337203685478: r = CInt(dA): out = out & "conv.CInt.d17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": r = CInt(vA): out = out & "conv.CInt.s0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "1.5": r = CInt(vA): out = out & "conv.CInt.s1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "2.5": r = CInt(vA): out = out & "conv.CInt.s2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "-0.5": r = CInt(vA): out = out & "conv.CInt.s3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = " 7 ": r = CInt(vA): out = out & "conv.CInt.s4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "1e3": r = CInt(vA): out = out & "conv.CInt.s5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "&H10": r = CInt(vA): out = out & "conv.CInt.s6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "abc": r = CInt(vA): out = out & "conv.CInt.s7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "": r = CInt(vA): out = out & "conv.CInt.s8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "32768": r = CInt(vA): out = out & "conv.CInt.s9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "99999999999": r = CInt(vA): out = out & "conv.CInt.s10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: r = CInt(vA): out = out & "conv.CInt.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: r = CInt(vA): out = out & "conv.CInt.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: r = CInt(vA): out = out & "conv.CInt.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: r = CInt(vA): out = out & "conv.CInt.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": r = CInt(vA): out = out & "conv.CInt.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: r = CInt(vA): out = out & "conv.CInt.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = -0.5: r = CLng(dA): out = out & "conv.CLng.d0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: r = CLng(dA): out = out & "conv.CLng.d1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.5: r = CLng(dA): out = out & "conv.CLng.d2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: r = CLng(dA): out = out & "conv.CLng.d3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 255.5: r = CLng(dA): out = out & "conv.CLng.d4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 256: r = CLng(dA): out = out & "conv.CLng.d5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 32767.5: r = CLng(dA): out = out & "conv.CLng.d6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 32768: r = CLng(dA): out = out & "conv.CLng.d7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2147483647.5: r = CLng(dA): out = out & "conv.CLng.d8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2147483648: r = CLng(dA): out = out & "conv.CLng.d9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 4.5E+18: r = CLng(dA): out = out & "conv.CLng.d10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 9.3E+18: r = CLng(dA): out = out & "conv.CLng.d11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = -1: r = CLng(dA): out = out & "conv.CLng.d12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+38: r = CLng(dA): out = out & "conv.CLng.d13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 3.5E+38: r = CLng(dA): out = out & "conv.CLng.d14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+300: r = CLng(dA): out = out & "conv.CLng.d15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 922337203685477.58: r = CLng(dA): out = out & "conv.CLng.d16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 922337203685478: r = CLng(dA): out = out & "conv.CLng.d17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": r = CLng(vA): out = out & "conv.CLng.s0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "1.5": r = CLng(vA): out = out & "conv.CLng.s1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "2.5": r = CLng(vA): out = out & "conv.CLng.s2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "-0.5": r = CLng(vA): out = out & "conv.CLng.s3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = " 7 ": r = CLng(vA): out = out & "conv.CLng.s4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "1e3": r = CLng(vA): out = out & "conv.CLng.s5=" & Rz(Err.Number, r) & vbLf
  P052 = out
End Function
Function P053() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: vA = "&H10": r = CLng(vA): out = out & "conv.CLng.s6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "abc": r = CLng(vA): out = out & "conv.CLng.s7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "": r = CLng(vA): out = out & "conv.CLng.s8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "32768": r = CLng(vA): out = out & "conv.CLng.s9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "99999999999": r = CLng(vA): out = out & "conv.CLng.s10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: r = CLng(vA): out = out & "conv.CLng.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: r = CLng(vA): out = out & "conv.CLng.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: r = CLng(vA): out = out & "conv.CLng.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: r = CLng(vA): out = out & "conv.CLng.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": r = CLng(vA): out = out & "conv.CLng.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: r = CLng(vA): out = out & "conv.CLng.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = -0.5: r = CLngLng(dA): out = out & "conv.CLngLng.d0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: r = CLngLng(dA): out = out & "conv.CLngLng.d1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.5: r = CLngLng(dA): out = out & "conv.CLngLng.d2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: r = CLngLng(dA): out = out & "conv.CLngLng.d3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 255.5: r = CLngLng(dA): out = out & "conv.CLngLng.d4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 256: r = CLngLng(dA): out = out & "conv.CLngLng.d5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 32767.5: r = CLngLng(dA): out = out & "conv.CLngLng.d6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 32768: r = CLngLng(dA): out = out & "conv.CLngLng.d7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2147483647.5: r = CLngLng(dA): out = out & "conv.CLngLng.d8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2147483648: r = CLngLng(dA): out = out & "conv.CLngLng.d9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 4.5E+18: r = CLngLng(dA): out = out & "conv.CLngLng.d10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 9.3E+18: r = CLngLng(dA): out = out & "conv.CLngLng.d11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = -1: r = CLngLng(dA): out = out & "conv.CLngLng.d12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+38: r = CLngLng(dA): out = out & "conv.CLngLng.d13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 3.5E+38: r = CLngLng(dA): out = out & "conv.CLngLng.d14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+300: r = CLngLng(dA): out = out & "conv.CLngLng.d15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 922337203685477.58: r = CLngLng(dA): out = out & "conv.CLngLng.d16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 922337203685478: r = CLngLng(dA): out = out & "conv.CLngLng.d17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": r = CLngLng(vA): out = out & "conv.CLngLng.s0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "1.5": r = CLngLng(vA): out = out & "conv.CLngLng.s1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "2.5": r = CLngLng(vA): out = out & "conv.CLngLng.s2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "-0.5": r = CLngLng(vA): out = out & "conv.CLngLng.s3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = " 7 ": r = CLngLng(vA): out = out & "conv.CLngLng.s4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "1e3": r = CLngLng(vA): out = out & "conv.CLngLng.s5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "&H10": r = CLngLng(vA): out = out & "conv.CLngLng.s6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "abc": r = CLngLng(vA): out = out & "conv.CLngLng.s7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "": r = CLngLng(vA): out = out & "conv.CLngLng.s8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "32768": r = CLngLng(vA): out = out & "conv.CLngLng.s9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "99999999999": r = CLngLng(vA): out = out & "conv.CLngLng.s10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: r = CLngLng(vA): out = out & "conv.CLngLng.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: r = CLngLng(vA): out = out & "conv.CLngLng.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: r = CLngLng(vA): out = out & "conv.CLngLng.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: r = CLngLng(vA): out = out & "conv.CLngLng.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": r = CLngLng(vA): out = out & "conv.CLngLng.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: r = CLngLng(vA): out = out & "conv.CLngLng.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = -0.5: r = CSng(dA): out = out & "conv.CSng.d0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: r = CSng(dA): out = out & "conv.CSng.d1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.5: r = CSng(dA): out = out & "conv.CSng.d2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: r = CSng(dA): out = out & "conv.CSng.d3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 255.5: r = CSng(dA): out = out & "conv.CSng.d4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 256: r = CSng(dA): out = out & "conv.CSng.d5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 32767.5: r = CSng(dA): out = out & "conv.CSng.d6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 32768: r = CSng(dA): out = out & "conv.CSng.d7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2147483647.5: r = CSng(dA): out = out & "conv.CSng.d8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2147483648: r = CSng(dA): out = out & "conv.CSng.d9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 4.5E+18: r = CSng(dA): out = out & "conv.CSng.d10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 9.3E+18: r = CSng(dA): out = out & "conv.CSng.d11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = -1: r = CSng(dA): out = out & "conv.CSng.d12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+38: r = CSng(dA): out = out & "conv.CSng.d13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 3.5E+38: r = CSng(dA): out = out & "conv.CSng.d14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+300: r = CSng(dA): out = out & "conv.CSng.d15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 922337203685477.58: r = CSng(dA): out = out & "conv.CSng.d16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 922337203685478: r = CSng(dA): out = out & "conv.CSng.d17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": r = CSng(vA): out = out & "conv.CSng.s0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "1.5": r = CSng(vA): out = out & "conv.CSng.s1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "2.5": r = CSng(vA): out = out & "conv.CSng.s2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "-0.5": r = CSng(vA): out = out & "conv.CSng.s3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = " 7 ": r = CSng(vA): out = out & "conv.CSng.s4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "1e3": r = CSng(vA): out = out & "conv.CSng.s5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "&H10": r = CSng(vA): out = out & "conv.CSng.s6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "abc": r = CSng(vA): out = out & "conv.CSng.s7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "": r = CSng(vA): out = out & "conv.CSng.s8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "32768": r = CSng(vA): out = out & "conv.CSng.s9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "99999999999": r = CSng(vA): out = out & "conv.CSng.s10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: r = CSng(vA): out = out & "conv.CSng.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: r = CSng(vA): out = out & "conv.CSng.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: r = CSng(vA): out = out & "conv.CSng.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: r = CSng(vA): out = out & "conv.CSng.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": r = CSng(vA): out = out & "conv.CSng.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: r = CSng(vA): out = out & "conv.CSng.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = -0.5: r = CDbl(dA): out = out & "conv.CDbl.d0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: r = CDbl(dA): out = out & "conv.CDbl.d1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.5: r = CDbl(dA): out = out & "conv.CDbl.d2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: r = CDbl(dA): out = out & "conv.CDbl.d3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 255.5: r = CDbl(dA): out = out & "conv.CDbl.d4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 256: r = CDbl(dA): out = out & "conv.CDbl.d5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 32767.5: r = CDbl(dA): out = out & "conv.CDbl.d6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 32768: r = CDbl(dA): out = out & "conv.CDbl.d7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2147483647.5: r = CDbl(dA): out = out & "conv.CDbl.d8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2147483648: r = CDbl(dA): out = out & "conv.CDbl.d9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 4.5E+18: r = CDbl(dA): out = out & "conv.CDbl.d10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 9.3E+18: r = CDbl(dA): out = out & "conv.CDbl.d11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = -1: r = CDbl(dA): out = out & "conv.CDbl.d12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+38: r = CDbl(dA): out = out & "conv.CDbl.d13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 3.5E+38: r = CDbl(dA): out = out & "conv.CDbl.d14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+300: r = CDbl(dA): out = out & "conv.CDbl.d15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 922337203685477.58: r = CDbl(dA): out = out & "conv.CDbl.d16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 922337203685478: r = CDbl(dA): out = out & "conv.CDbl.d17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": r = CDbl(vA): out = out & "conv.CDbl.s0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "1.5": r = CDbl(vA): out = out & "conv.CDbl.s1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "2.5": r = CDbl(vA): out = out & "conv.CDbl.s2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "-0.5": r = CDbl(vA): out = out & "conv.CDbl.s3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = " 7 ": r = CDbl(vA): out = out & "conv.CDbl.s4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "1e3": r = CDbl(vA): out = out & "conv.CDbl.s5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "&H10": r = CDbl(vA): out = out & "conv.CDbl.s6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "abc": r = CDbl(vA): out = out & "conv.CDbl.s7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "": r = CDbl(vA): out = out & "conv.CDbl.s8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "32768": r = CDbl(vA): out = out & "conv.CDbl.s9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "99999999999": r = CDbl(vA): out = out & "conv.CDbl.s10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: r = CDbl(vA): out = out & "conv.CDbl.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: r = CDbl(vA): out = out & "conv.CDbl.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: r = CDbl(vA): out = out & "conv.CDbl.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: r = CDbl(vA): out = out & "conv.CDbl.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": r = CDbl(vA): out = out & "conv.CDbl.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: r = CDbl(vA): out = out & "conv.CDbl.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = -0.5: r = CCur(dA): out = out & "conv.CCur.d0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: r = CCur(dA): out = out & "conv.CCur.d1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.5: r = CCur(dA): out = out & "conv.CCur.d2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: r = CCur(dA): out = out & "conv.CCur.d3=" & Rz(Err.Number, r) & vbLf
  P053 = out
End Function
Function P054() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: dA = 255.5: r = CCur(dA): out = out & "conv.CCur.d4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 256: r = CCur(dA): out = out & "conv.CCur.d5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 32767.5: r = CCur(dA): out = out & "conv.CCur.d6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 32768: r = CCur(dA): out = out & "conv.CCur.d7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2147483647.5: r = CCur(dA): out = out & "conv.CCur.d8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2147483648: r = CCur(dA): out = out & "conv.CCur.d9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 4.5E+18: r = CCur(dA): out = out & "conv.CCur.d10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 9.3E+18: r = CCur(dA): out = out & "conv.CCur.d11=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = -1: r = CCur(dA): out = out & "conv.CCur.d12=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+38: r = CCur(dA): out = out & "conv.CCur.d13=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 3.5E+38: r = CCur(dA): out = out & "conv.CCur.d14=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+300: r = CCur(dA): out = out & "conv.CCur.d15=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 922337203685477.58: r = CCur(dA): out = out & "conv.CCur.d16=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 922337203685478: r = CCur(dA): out = out & "conv.CCur.d17=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": r = CCur(vA): out = out & "conv.CCur.s0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "1.5": r = CCur(vA): out = out & "conv.CCur.s1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "2.5": r = CCur(vA): out = out & "conv.CCur.s2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "-0.5": r = CCur(vA): out = out & "conv.CCur.s3=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = " 7 ": r = CCur(vA): out = out & "conv.CCur.s4=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "1e3": r = CCur(vA): out = out & "conv.CCur.s5=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "&H10": r = CCur(vA): out = out & "conv.CCur.s6=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "abc": r = CCur(vA): out = out & "conv.CCur.s7=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "": r = CCur(vA): out = out & "conv.CCur.s8=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "32768": r = CCur(vA): out = out & "conv.CCur.s9=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "99999999999": r = CCur(vA): out = out & "conv.CCur.s10=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 32767: r = CCur(vA): out = out & "conv.CCur.vInt=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2147483647: r = CCur(vA): out = out & "conv.CCur.vLng=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 2.5: r = CCur(vA): out = out & "conv.CCur.vDbl=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = 1.5@: r = CCur(vA): out = out & "conv.CCur.vCur=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = "12": r = CCur(vA): out = out & "conv.CCur.vStr=" & Rz(Err.Number, r) & vbLf
  Err.Clear: vA = Empty: r = CCur(vA): out = out & "conv.CCur.vEmp=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: bB = 0: r = (bA = bB): out = out & "cmp.Byte0.eq.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: bB = 0: r = (bA < bB): out = out & "cmp.Byte0.lt.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: bB = 200: r = (bA = bB): out = out & "cmp.Byte0.eq.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: bB = 200: r = (bA < bB): out = out & "cmp.Byte0.lt.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: bB = 255: r = (bA = bB): out = out & "cmp.Byte0.eq.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: bB = 255: r = (bA < bB): out = out & "cmp.Byte0.lt.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: bB = 0: r = (bA = bB): out = out & "cmp.Byte1.eq.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: bB = 0: r = (bA < bB): out = out & "cmp.Byte1.lt.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: bB = 200: r = (bA = bB): out = out & "cmp.Byte1.eq.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: bB = 200: r = (bA < bB): out = out & "cmp.Byte1.lt.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: bB = 255: r = (bA = bB): out = out & "cmp.Byte1.eq.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: bB = 255: r = (bA < bB): out = out & "cmp.Byte1.lt.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: bB = 0: r = (bA = bB): out = out & "cmp.Byte2.eq.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: bB = 0: r = (bA < bB): out = out & "cmp.Byte2.lt.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: bB = 200: r = (bA = bB): out = out & "cmp.Byte2.eq.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: bB = 200: r = (bA < bB): out = out & "cmp.Byte2.lt.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: bB = 255: r = (bA = bB): out = out & "cmp.Byte2.eq.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: bB = 255: r = (bA < bB): out = out & "cmp.Byte2.lt.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: iB = -32768: r = (bA = iB): out = out & "cmp.Byte0.eq.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: iB = -32768: r = (bA < iB): out = out & "cmp.Byte0.lt.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: iB = 1: r = (bA = iB): out = out & "cmp.Byte0.eq.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: iB = 1: r = (bA < iB): out = out & "cmp.Byte0.lt.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: iB = 32767: r = (bA = iB): out = out & "cmp.Byte0.eq.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: iB = 32767: r = (bA < iB): out = out & "cmp.Byte0.lt.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: iB = -32768: r = (bA = iB): out = out & "cmp.Byte1.eq.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: iB = -32768: r = (bA < iB): out = out & "cmp.Byte1.lt.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: iB = 1: r = (bA = iB): out = out & "cmp.Byte1.eq.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: iB = 1: r = (bA < iB): out = out & "cmp.Byte1.lt.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: iB = 32767: r = (bA = iB): out = out & "cmp.Byte1.eq.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: iB = 32767: r = (bA < iB): out = out & "cmp.Byte1.lt.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: iB = -32768: r = (bA = iB): out = out & "cmp.Byte2.eq.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: iB = -32768: r = (bA < iB): out = out & "cmp.Byte2.lt.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: iB = 1: r = (bA = iB): out = out & "cmp.Byte2.eq.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: iB = 1: r = (bA < iB): out = out & "cmp.Byte2.lt.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: iB = 32767: r = (bA = iB): out = out & "cmp.Byte2.eq.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: iB = 32767: r = (bA < iB): out = out & "cmp.Byte2.lt.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: lB = -2147483648: r = (bA = lB): out = out & "cmp.Byte0.eq.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: lB = -2147483648: r = (bA < lB): out = out & "cmp.Byte0.lt.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: lB = 65536: r = (bA = lB): out = out & "cmp.Byte0.eq.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: lB = 65536: r = (bA < lB): out = out & "cmp.Byte0.lt.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: lB = 2147483647: r = (bA = lB): out = out & "cmp.Byte0.eq.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: lB = 2147483647: r = (bA < lB): out = out & "cmp.Byte0.lt.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: lB = -2147483648: r = (bA = lB): out = out & "cmp.Byte1.eq.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: lB = -2147483648: r = (bA < lB): out = out & "cmp.Byte1.lt.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: lB = 65536: r = (bA = lB): out = out & "cmp.Byte1.eq.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: lB = 65536: r = (bA < lB): out = out & "cmp.Byte1.lt.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: lB = 2147483647: r = (bA = lB): out = out & "cmp.Byte1.eq.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: lB = 2147483647: r = (bA < lB): out = out & "cmp.Byte1.lt.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: lB = -2147483648: r = (bA = lB): out = out & "cmp.Byte2.eq.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: lB = -2147483648: r = (bA < lB): out = out & "cmp.Byte2.lt.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: lB = 65536: r = (bA = lB): out = out & "cmp.Byte2.eq.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: lB = 65536: r = (bA < lB): out = out & "cmp.Byte2.lt.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: lB = 2147483647: r = (bA = lB): out = out & "cmp.Byte2.eq.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: lB = 2147483647: r = (bA < lB): out = out & "cmp.Byte2.lt.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: qB = -9223372036854775807^: r = (bA = qB): out = out & "cmp.Byte0.eq.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: qB = -9223372036854775807^: r = (bA < qB): out = out & "cmp.Byte0.lt.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: qB = 3: r = (bA = qB): out = out & "cmp.Byte0.eq.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: qB = 3: r = (bA < qB): out = out & "cmp.Byte0.lt.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: qB = 9223372036854775807^: r = (bA = qB): out = out & "cmp.Byte0.eq.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: qB = 9223372036854775807^: r = (bA < qB): out = out & "cmp.Byte0.lt.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: qB = -9223372036854775807^: r = (bA = qB): out = out & "cmp.Byte1.eq.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: qB = -9223372036854775807^: r = (bA < qB): out = out & "cmp.Byte1.lt.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: qB = 3: r = (bA = qB): out = out & "cmp.Byte1.eq.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: qB = 3: r = (bA < qB): out = out & "cmp.Byte1.lt.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: qB = 9223372036854775807^: r = (bA = qB): out = out & "cmp.Byte1.eq.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: qB = 9223372036854775807^: r = (bA < qB): out = out & "cmp.Byte1.lt.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: qB = -9223372036854775807^: r = (bA = qB): out = out & "cmp.Byte2.eq.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: qB = -9223372036854775807^: r = (bA < qB): out = out & "cmp.Byte2.lt.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: qB = 3: r = (bA = qB): out = out & "cmp.Byte2.eq.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: qB = 3: r = (bA < qB): out = out & "cmp.Byte2.lt.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: qB = 9223372036854775807^: r = (bA = qB): out = out & "cmp.Byte2.eq.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: qB = 9223372036854775807^: r = (bA < qB): out = out & "cmp.Byte2.lt.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: sB = 0.5: r = (bA = sB): out = out & "cmp.Byte0.eq.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: sB = 0.5: r = (bA < sB): out = out & "cmp.Byte0.lt.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: sB = 16777216: r = (bA = sB): out = out & "cmp.Byte0.eq.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: sB = 16777216: r = (bA < sB): out = out & "cmp.Byte0.lt.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: sB = 3.0E+38: r = (bA = sB): out = out & "cmp.Byte0.eq.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: sB = 3.0E+38: r = (bA < sB): out = out & "cmp.Byte0.lt.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: sB = 0.5: r = (bA = sB): out = out & "cmp.Byte1.eq.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: sB = 0.5: r = (bA < sB): out = out & "cmp.Byte1.lt.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: sB = 16777216: r = (bA = sB): out = out & "cmp.Byte1.eq.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: sB = 16777216: r = (bA < sB): out = out & "cmp.Byte1.lt.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: sB = 3.0E+38: r = (bA = sB): out = out & "cmp.Byte1.eq.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: sB = 3.0E+38: r = (bA < sB): out = out & "cmp.Byte1.lt.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: sB = 0.5: r = (bA = sB): out = out & "cmp.Byte2.eq.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: sB = 0.5: r = (bA < sB): out = out & "cmp.Byte2.lt.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: sB = 16777216: r = (bA = sB): out = out & "cmp.Byte2.eq.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: sB = 16777216: r = (bA < sB): out = out & "cmp.Byte2.lt.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: sB = 3.0E+38: r = (bA = sB): out = out & "cmp.Byte2.eq.Single2=" & Rz(Err.Number, r) & vbLf
  P054 = out
End Function
Function P055() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: bA = 255: sB = 3.0E+38: r = (bA < sB): out = out & "cmp.Byte2.lt.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: dB = 0.5: r = (bA = dB): out = out & "cmp.Byte0.eq.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: dB = 0.5: r = (bA < dB): out = out & "cmp.Byte0.lt.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: dB = 2.5: r = (bA = dB): out = out & "cmp.Byte0.eq.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: dB = 2.5: r = (bA < dB): out = out & "cmp.Byte0.lt.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: dB = 1.0E+308: r = (bA = dB): out = out & "cmp.Byte0.eq.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: dB = 1.0E+308: r = (bA < dB): out = out & "cmp.Byte0.lt.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: dB = 0.5: r = (bA = dB): out = out & "cmp.Byte1.eq.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: dB = 0.5: r = (bA < dB): out = out & "cmp.Byte1.lt.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: dB = 2.5: r = (bA = dB): out = out & "cmp.Byte1.eq.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: dB = 2.5: r = (bA < dB): out = out & "cmp.Byte1.lt.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: dB = 1.0E+308: r = (bA = dB): out = out & "cmp.Byte1.eq.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: dB = 1.0E+308: r = (bA < dB): out = out & "cmp.Byte1.lt.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: dB = 0.5: r = (bA = dB): out = out & "cmp.Byte2.eq.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: dB = 0.5: r = (bA < dB): out = out & "cmp.Byte2.lt.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: dB = 2.5: r = (bA = dB): out = out & "cmp.Byte2.eq.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: dB = 2.5: r = (bA < dB): out = out & "cmp.Byte2.lt.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: dB = 1.0E+308: r = (bA = dB): out = out & "cmp.Byte2.eq.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: dB = 1.0E+308: r = (bA < dB): out = out & "cmp.Byte2.lt.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: cB = -0.5@: r = (bA = cB): out = out & "cmp.Byte0.eq.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: cB = -0.5@: r = (bA < cB): out = out & "cmp.Byte0.lt.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: cB = 1.5@: r = (bA = cB): out = out & "cmp.Byte0.eq.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: cB = 1.5@: r = (bA < cB): out = out & "cmp.Byte0.lt.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: cB = 922337203685477.5807@: r = (bA = cB): out = out & "cmp.Byte0.eq.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 0: cB = 922337203685477.5807@: r = (bA < cB): out = out & "cmp.Byte0.lt.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: cB = -0.5@: r = (bA = cB): out = out & "cmp.Byte1.eq.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: cB = -0.5@: r = (bA < cB): out = out & "cmp.Byte1.lt.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: cB = 1.5@: r = (bA = cB): out = out & "cmp.Byte1.eq.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: cB = 1.5@: r = (bA < cB): out = out & "cmp.Byte1.lt.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: cB = 922337203685477.5807@: r = (bA = cB): out = out & "cmp.Byte1.eq.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 200: cB = 922337203685477.5807@: r = (bA < cB): out = out & "cmp.Byte1.lt.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: cB = -0.5@: r = (bA = cB): out = out & "cmp.Byte2.eq.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: cB = -0.5@: r = (bA < cB): out = out & "cmp.Byte2.lt.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: cB = 1.5@: r = (bA = cB): out = out & "cmp.Byte2.eq.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: cB = 1.5@: r = (bA < cB): out = out & "cmp.Byte2.lt.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: cB = 922337203685477.5807@: r = (bA = cB): out = out & "cmp.Byte2.eq.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: bA = 255: cB = 922337203685477.5807@: r = (bA < cB): out = out & "cmp.Byte2.lt.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: bB = 0: r = (iA = bB): out = out & "cmp.Integer0.eq.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: bB = 0: r = (iA < bB): out = out & "cmp.Integer0.lt.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: bB = 200: r = (iA = bB): out = out & "cmp.Integer0.eq.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: bB = 200: r = (iA < bB): out = out & "cmp.Integer0.lt.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: bB = 255: r = (iA = bB): out = out & "cmp.Integer0.eq.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: bB = 255: r = (iA < bB): out = out & "cmp.Integer0.lt.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: bB = 0: r = (iA = bB): out = out & "cmp.Integer1.eq.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: bB = 0: r = (iA < bB): out = out & "cmp.Integer1.lt.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: bB = 200: r = (iA = bB): out = out & "cmp.Integer1.eq.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: bB = 200: r = (iA < bB): out = out & "cmp.Integer1.lt.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: bB = 255: r = (iA = bB): out = out & "cmp.Integer1.eq.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: bB = 255: r = (iA < bB): out = out & "cmp.Integer1.lt.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: bB = 0: r = (iA = bB): out = out & "cmp.Integer2.eq.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: bB = 0: r = (iA < bB): out = out & "cmp.Integer2.lt.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: bB = 200: r = (iA = bB): out = out & "cmp.Integer2.eq.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: bB = 200: r = (iA < bB): out = out & "cmp.Integer2.lt.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: bB = 255: r = (iA = bB): out = out & "cmp.Integer2.eq.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: bB = 255: r = (iA < bB): out = out & "cmp.Integer2.lt.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: iB = -32768: r = (iA = iB): out = out & "cmp.Integer0.eq.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: iB = -32768: r = (iA < iB): out = out & "cmp.Integer0.lt.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: iB = 1: r = (iA = iB): out = out & "cmp.Integer0.eq.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: iB = 1: r = (iA < iB): out = out & "cmp.Integer0.lt.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: iB = 32767: r = (iA = iB): out = out & "cmp.Integer0.eq.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: iB = 32767: r = (iA < iB): out = out & "cmp.Integer0.lt.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: iB = -32768: r = (iA = iB): out = out & "cmp.Integer1.eq.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: iB = -32768: r = (iA < iB): out = out & "cmp.Integer1.lt.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: iB = 1: r = (iA = iB): out = out & "cmp.Integer1.eq.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: iB = 1: r = (iA < iB): out = out & "cmp.Integer1.lt.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: iB = 32767: r = (iA = iB): out = out & "cmp.Integer1.eq.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: iB = 32767: r = (iA < iB): out = out & "cmp.Integer1.lt.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: iB = -32768: r = (iA = iB): out = out & "cmp.Integer2.eq.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: iB = -32768: r = (iA < iB): out = out & "cmp.Integer2.lt.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: iB = 1: r = (iA = iB): out = out & "cmp.Integer2.eq.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: iB = 1: r = (iA < iB): out = out & "cmp.Integer2.lt.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: iB = 32767: r = (iA = iB): out = out & "cmp.Integer2.eq.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: iB = 32767: r = (iA < iB): out = out & "cmp.Integer2.lt.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: lB = -2147483648: r = (iA = lB): out = out & "cmp.Integer0.eq.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: lB = -2147483648: r = (iA < lB): out = out & "cmp.Integer0.lt.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: lB = 65536: r = (iA = lB): out = out & "cmp.Integer0.eq.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: lB = 65536: r = (iA < lB): out = out & "cmp.Integer0.lt.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: lB = 2147483647: r = (iA = lB): out = out & "cmp.Integer0.eq.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: lB = 2147483647: r = (iA < lB): out = out & "cmp.Integer0.lt.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: lB = -2147483648: r = (iA = lB): out = out & "cmp.Integer1.eq.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: lB = -2147483648: r = (iA < lB): out = out & "cmp.Integer1.lt.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: lB = 65536: r = (iA = lB): out = out & "cmp.Integer1.eq.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: lB = 65536: r = (iA < lB): out = out & "cmp.Integer1.lt.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: lB = 2147483647: r = (iA = lB): out = out & "cmp.Integer1.eq.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: lB = 2147483647: r = (iA < lB): out = out & "cmp.Integer1.lt.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: lB = -2147483648: r = (iA = lB): out = out & "cmp.Integer2.eq.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: lB = -2147483648: r = (iA < lB): out = out & "cmp.Integer2.lt.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: lB = 65536: r = (iA = lB): out = out & "cmp.Integer2.eq.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: lB = 65536: r = (iA < lB): out = out & "cmp.Integer2.lt.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: lB = 2147483647: r = (iA = lB): out = out & "cmp.Integer2.eq.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: lB = 2147483647: r = (iA < lB): out = out & "cmp.Integer2.lt.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: qB = -9223372036854775807^: r = (iA = qB): out = out & "cmp.Integer0.eq.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: qB = -9223372036854775807^: r = (iA < qB): out = out & "cmp.Integer0.lt.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: qB = 3: r = (iA = qB): out = out & "cmp.Integer0.eq.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: qB = 3: r = (iA < qB): out = out & "cmp.Integer0.lt.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: qB = 9223372036854775807^: r = (iA = qB): out = out & "cmp.Integer0.eq.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: qB = 9223372036854775807^: r = (iA < qB): out = out & "cmp.Integer0.lt.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: qB = -9223372036854775807^: r = (iA = qB): out = out & "cmp.Integer1.eq.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: qB = -9223372036854775807^: r = (iA < qB): out = out & "cmp.Integer1.lt.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: qB = 3: r = (iA = qB): out = out & "cmp.Integer1.eq.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: qB = 3: r = (iA < qB): out = out & "cmp.Integer1.lt.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: qB = 9223372036854775807^: r = (iA = qB): out = out & "cmp.Integer1.eq.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: qB = 9223372036854775807^: r = (iA < qB): out = out & "cmp.Integer1.lt.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: qB = -9223372036854775807^: r = (iA = qB): out = out & "cmp.Integer2.eq.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: qB = -9223372036854775807^: r = (iA < qB): out = out & "cmp.Integer2.lt.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: qB = 3: r = (iA = qB): out = out & "cmp.Integer2.eq.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: qB = 3: r = (iA < qB): out = out & "cmp.Integer2.lt.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: qB = 9223372036854775807^: r = (iA = qB): out = out & "cmp.Integer2.eq.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: qB = 9223372036854775807^: r = (iA < qB): out = out & "cmp.Integer2.lt.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: sB = 0.5: r = (iA = sB): out = out & "cmp.Integer0.eq.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: sB = 0.5: r = (iA < sB): out = out & "cmp.Integer0.lt.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: sB = 16777216: r = (iA = sB): out = out & "cmp.Integer0.eq.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: sB = 16777216: r = (iA < sB): out = out & "cmp.Integer0.lt.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: sB = 3.0E+38: r = (iA = sB): out = out & "cmp.Integer0.eq.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: sB = 3.0E+38: r = (iA < sB): out = out & "cmp.Integer0.lt.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: sB = 0.5: r = (iA = sB): out = out & "cmp.Integer1.eq.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: sB = 0.5: r = (iA < sB): out = out & "cmp.Integer1.lt.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: sB = 16777216: r = (iA = sB): out = out & "cmp.Integer1.eq.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: sB = 16777216: r = (iA < sB): out = out & "cmp.Integer1.lt.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: sB = 3.0E+38: r = (iA = sB): out = out & "cmp.Integer1.eq.Single2=" & Rz(Err.Number, r) & vbLf
  P055 = out
End Function
Function P056() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: iA = 1: sB = 3.0E+38: r = (iA < sB): out = out & "cmp.Integer1.lt.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: sB = 0.5: r = (iA = sB): out = out & "cmp.Integer2.eq.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: sB = 0.5: r = (iA < sB): out = out & "cmp.Integer2.lt.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: sB = 16777216: r = (iA = sB): out = out & "cmp.Integer2.eq.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: sB = 16777216: r = (iA < sB): out = out & "cmp.Integer2.lt.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: sB = 3.0E+38: r = (iA = sB): out = out & "cmp.Integer2.eq.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: sB = 3.0E+38: r = (iA < sB): out = out & "cmp.Integer2.lt.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: dB = 0.5: r = (iA = dB): out = out & "cmp.Integer0.eq.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: dB = 0.5: r = (iA < dB): out = out & "cmp.Integer0.lt.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: dB = 2.5: r = (iA = dB): out = out & "cmp.Integer0.eq.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: dB = 2.5: r = (iA < dB): out = out & "cmp.Integer0.lt.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: dB = 1.0E+308: r = (iA = dB): out = out & "cmp.Integer0.eq.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: dB = 1.0E+308: r = (iA < dB): out = out & "cmp.Integer0.lt.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: dB = 0.5: r = (iA = dB): out = out & "cmp.Integer1.eq.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: dB = 0.5: r = (iA < dB): out = out & "cmp.Integer1.lt.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: dB = 2.5: r = (iA = dB): out = out & "cmp.Integer1.eq.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: dB = 2.5: r = (iA < dB): out = out & "cmp.Integer1.lt.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: dB = 1.0E+308: r = (iA = dB): out = out & "cmp.Integer1.eq.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: dB = 1.0E+308: r = (iA < dB): out = out & "cmp.Integer1.lt.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: dB = 0.5: r = (iA = dB): out = out & "cmp.Integer2.eq.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: dB = 0.5: r = (iA < dB): out = out & "cmp.Integer2.lt.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: dB = 2.5: r = (iA = dB): out = out & "cmp.Integer2.eq.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: dB = 2.5: r = (iA < dB): out = out & "cmp.Integer2.lt.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: dB = 1.0E+308: r = (iA = dB): out = out & "cmp.Integer2.eq.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: dB = 1.0E+308: r = (iA < dB): out = out & "cmp.Integer2.lt.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: cB = -0.5@: r = (iA = cB): out = out & "cmp.Integer0.eq.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: cB = -0.5@: r = (iA < cB): out = out & "cmp.Integer0.lt.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: cB = 1.5@: r = (iA = cB): out = out & "cmp.Integer0.eq.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: cB = 1.5@: r = (iA < cB): out = out & "cmp.Integer0.lt.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: cB = 922337203685477.5807@: r = (iA = cB): out = out & "cmp.Integer0.eq.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = -32768: cB = 922337203685477.5807@: r = (iA < cB): out = out & "cmp.Integer0.lt.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: cB = -0.5@: r = (iA = cB): out = out & "cmp.Integer1.eq.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: cB = -0.5@: r = (iA < cB): out = out & "cmp.Integer1.lt.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: cB = 1.5@: r = (iA = cB): out = out & "cmp.Integer1.eq.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: cB = 1.5@: r = (iA < cB): out = out & "cmp.Integer1.lt.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: cB = 922337203685477.5807@: r = (iA = cB): out = out & "cmp.Integer1.eq.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 1: cB = 922337203685477.5807@: r = (iA < cB): out = out & "cmp.Integer1.lt.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: cB = -0.5@: r = (iA = cB): out = out & "cmp.Integer2.eq.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: cB = -0.5@: r = (iA < cB): out = out & "cmp.Integer2.lt.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: cB = 1.5@: r = (iA = cB): out = out & "cmp.Integer2.eq.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: cB = 1.5@: r = (iA < cB): out = out & "cmp.Integer2.lt.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: cB = 922337203685477.5807@: r = (iA = cB): out = out & "cmp.Integer2.eq.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: iA = 32767: cB = 922337203685477.5807@: r = (iA < cB): out = out & "cmp.Integer2.lt.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: bB = 0: r = (lA = bB): out = out & "cmp.Long0.eq.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: bB = 0: r = (lA < bB): out = out & "cmp.Long0.lt.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: bB = 200: r = (lA = bB): out = out & "cmp.Long0.eq.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: bB = 200: r = (lA < bB): out = out & "cmp.Long0.lt.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: bB = 255: r = (lA = bB): out = out & "cmp.Long0.eq.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: bB = 255: r = (lA < bB): out = out & "cmp.Long0.lt.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: bB = 0: r = (lA = bB): out = out & "cmp.Long1.eq.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: bB = 0: r = (lA < bB): out = out & "cmp.Long1.lt.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: bB = 200: r = (lA = bB): out = out & "cmp.Long1.eq.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: bB = 200: r = (lA < bB): out = out & "cmp.Long1.lt.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: bB = 255: r = (lA = bB): out = out & "cmp.Long1.eq.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: bB = 255: r = (lA < bB): out = out & "cmp.Long1.lt.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: bB = 0: r = (lA = bB): out = out & "cmp.Long2.eq.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: bB = 0: r = (lA < bB): out = out & "cmp.Long2.lt.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: bB = 200: r = (lA = bB): out = out & "cmp.Long2.eq.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: bB = 200: r = (lA < bB): out = out & "cmp.Long2.lt.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: bB = 255: r = (lA = bB): out = out & "cmp.Long2.eq.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: bB = 255: r = (lA < bB): out = out & "cmp.Long2.lt.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: iB = -32768: r = (lA = iB): out = out & "cmp.Long0.eq.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: iB = -32768: r = (lA < iB): out = out & "cmp.Long0.lt.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: iB = 1: r = (lA = iB): out = out & "cmp.Long0.eq.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: iB = 1: r = (lA < iB): out = out & "cmp.Long0.lt.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: iB = 32767: r = (lA = iB): out = out & "cmp.Long0.eq.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: iB = 32767: r = (lA < iB): out = out & "cmp.Long0.lt.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: iB = -32768: r = (lA = iB): out = out & "cmp.Long1.eq.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: iB = -32768: r = (lA < iB): out = out & "cmp.Long1.lt.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: iB = 1: r = (lA = iB): out = out & "cmp.Long1.eq.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: iB = 1: r = (lA < iB): out = out & "cmp.Long1.lt.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: iB = 32767: r = (lA = iB): out = out & "cmp.Long1.eq.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: iB = 32767: r = (lA < iB): out = out & "cmp.Long1.lt.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: iB = -32768: r = (lA = iB): out = out & "cmp.Long2.eq.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: iB = -32768: r = (lA < iB): out = out & "cmp.Long2.lt.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: iB = 1: r = (lA = iB): out = out & "cmp.Long2.eq.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: iB = 1: r = (lA < iB): out = out & "cmp.Long2.lt.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: iB = 32767: r = (lA = iB): out = out & "cmp.Long2.eq.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: iB = 32767: r = (lA < iB): out = out & "cmp.Long2.lt.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: lB = -2147483648: r = (lA = lB): out = out & "cmp.Long0.eq.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: lB = -2147483648: r = (lA < lB): out = out & "cmp.Long0.lt.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: lB = 65536: r = (lA = lB): out = out & "cmp.Long0.eq.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: lB = 65536: r = (lA < lB): out = out & "cmp.Long0.lt.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: lB = 2147483647: r = (lA = lB): out = out & "cmp.Long0.eq.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: lB = 2147483647: r = (lA < lB): out = out & "cmp.Long0.lt.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: lB = -2147483648: r = (lA = lB): out = out & "cmp.Long1.eq.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: lB = -2147483648: r = (lA < lB): out = out & "cmp.Long1.lt.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: lB = 65536: r = (lA = lB): out = out & "cmp.Long1.eq.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: lB = 65536: r = (lA < lB): out = out & "cmp.Long1.lt.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: lB = 2147483647: r = (lA = lB): out = out & "cmp.Long1.eq.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: lB = 2147483647: r = (lA < lB): out = out & "cmp.Long1.lt.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: lB = -2147483648: r = (lA = lB): out = out & "cmp.Long2.eq.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: lB = -2147483648: r = (lA < lB): out = out & "cmp.Long2.lt.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: lB = 65536: r = (lA = lB): out = out & "cmp.Long2.eq.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: lB = 65536: r = (lA < lB): out = out & "cmp.Long2.lt.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: lB = 2147483647: r = (lA = lB): out = out & "cmp.Long2.eq.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: lB = 2147483647: r = (lA < lB): out = out & "cmp.Long2.lt.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: qB = -9223372036854775807^: r = (lA = qB): out = out & "cmp.Long0.eq.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: qB = -9223372036854775807^: r = (lA < qB): out = out & "cmp.Long0.lt.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: qB = 3: r = (lA = qB): out = out & "cmp.Long0.eq.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: qB = 3: r = (lA < qB): out = out & "cmp.Long0.lt.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: qB = 9223372036854775807^: r = (lA = qB): out = out & "cmp.Long0.eq.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: qB = 9223372036854775807^: r = (lA < qB): out = out & "cmp.Long0.lt.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: qB = -9223372036854775807^: r = (lA = qB): out = out & "cmp.Long1.eq.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: qB = -9223372036854775807^: r = (lA < qB): out = out & "cmp.Long1.lt.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: qB = 3: r = (lA = qB): out = out & "cmp.Long1.eq.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: qB = 3: r = (lA < qB): out = out & "cmp.Long1.lt.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: qB = 9223372036854775807^: r = (lA = qB): out = out & "cmp.Long1.eq.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: qB = 9223372036854775807^: r = (lA < qB): out = out & "cmp.Long1.lt.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: qB = -9223372036854775807^: r = (lA = qB): out = out & "cmp.Long2.eq.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: qB = -9223372036854775807^: r = (lA < qB): out = out & "cmp.Long2.lt.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: qB = 3: r = (lA = qB): out = out & "cmp.Long2.eq.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: qB = 3: r = (lA < qB): out = out & "cmp.Long2.lt.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: qB = 9223372036854775807^: r = (lA = qB): out = out & "cmp.Long2.eq.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: qB = 9223372036854775807^: r = (lA < qB): out = out & "cmp.Long2.lt.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: sB = 0.5: r = (lA = sB): out = out & "cmp.Long0.eq.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: sB = 0.5: r = (lA < sB): out = out & "cmp.Long0.lt.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: sB = 16777216: r = (lA = sB): out = out & "cmp.Long0.eq.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: sB = 16777216: r = (lA < sB): out = out & "cmp.Long0.lt.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: sB = 3.0E+38: r = (lA = sB): out = out & "cmp.Long0.eq.Single2=" & Rz(Err.Number, r) & vbLf
  P056 = out
End Function
Function P057() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: lA = -2147483648: sB = 3.0E+38: r = (lA < sB): out = out & "cmp.Long0.lt.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: sB = 0.5: r = (lA = sB): out = out & "cmp.Long1.eq.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: sB = 0.5: r = (lA < sB): out = out & "cmp.Long1.lt.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: sB = 16777216: r = (lA = sB): out = out & "cmp.Long1.eq.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: sB = 16777216: r = (lA < sB): out = out & "cmp.Long1.lt.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: sB = 3.0E+38: r = (lA = sB): out = out & "cmp.Long1.eq.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: sB = 3.0E+38: r = (lA < sB): out = out & "cmp.Long1.lt.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: sB = 0.5: r = (lA = sB): out = out & "cmp.Long2.eq.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: sB = 0.5: r = (lA < sB): out = out & "cmp.Long2.lt.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: sB = 16777216: r = (lA = sB): out = out & "cmp.Long2.eq.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: sB = 16777216: r = (lA < sB): out = out & "cmp.Long2.lt.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: sB = 3.0E+38: r = (lA = sB): out = out & "cmp.Long2.eq.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: sB = 3.0E+38: r = (lA < sB): out = out & "cmp.Long2.lt.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: dB = 0.5: r = (lA = dB): out = out & "cmp.Long0.eq.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: dB = 0.5: r = (lA < dB): out = out & "cmp.Long0.lt.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: dB = 2.5: r = (lA = dB): out = out & "cmp.Long0.eq.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: dB = 2.5: r = (lA < dB): out = out & "cmp.Long0.lt.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: dB = 1.0E+308: r = (lA = dB): out = out & "cmp.Long0.eq.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: dB = 1.0E+308: r = (lA < dB): out = out & "cmp.Long0.lt.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: dB = 0.5: r = (lA = dB): out = out & "cmp.Long1.eq.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: dB = 0.5: r = (lA < dB): out = out & "cmp.Long1.lt.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: dB = 2.5: r = (lA = dB): out = out & "cmp.Long1.eq.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: dB = 2.5: r = (lA < dB): out = out & "cmp.Long1.lt.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: dB = 1.0E+308: r = (lA = dB): out = out & "cmp.Long1.eq.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: dB = 1.0E+308: r = (lA < dB): out = out & "cmp.Long1.lt.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: dB = 0.5: r = (lA = dB): out = out & "cmp.Long2.eq.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: dB = 0.5: r = (lA < dB): out = out & "cmp.Long2.lt.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: dB = 2.5: r = (lA = dB): out = out & "cmp.Long2.eq.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: dB = 2.5: r = (lA < dB): out = out & "cmp.Long2.lt.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: dB = 1.0E+308: r = (lA = dB): out = out & "cmp.Long2.eq.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: dB = 1.0E+308: r = (lA < dB): out = out & "cmp.Long2.lt.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: cB = -0.5@: r = (lA = cB): out = out & "cmp.Long0.eq.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: cB = -0.5@: r = (lA < cB): out = out & "cmp.Long0.lt.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: cB = 1.5@: r = (lA = cB): out = out & "cmp.Long0.eq.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: cB = 1.5@: r = (lA < cB): out = out & "cmp.Long0.lt.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: cB = 922337203685477.5807@: r = (lA = cB): out = out & "cmp.Long0.eq.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = -2147483648: cB = 922337203685477.5807@: r = (lA < cB): out = out & "cmp.Long0.lt.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: cB = -0.5@: r = (lA = cB): out = out & "cmp.Long1.eq.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: cB = -0.5@: r = (lA < cB): out = out & "cmp.Long1.lt.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: cB = 1.5@: r = (lA = cB): out = out & "cmp.Long1.eq.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: cB = 1.5@: r = (lA < cB): out = out & "cmp.Long1.lt.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: cB = 922337203685477.5807@: r = (lA = cB): out = out & "cmp.Long1.eq.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 65536: cB = 922337203685477.5807@: r = (lA < cB): out = out & "cmp.Long1.lt.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: cB = -0.5@: r = (lA = cB): out = out & "cmp.Long2.eq.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: cB = -0.5@: r = (lA < cB): out = out & "cmp.Long2.lt.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: cB = 1.5@: r = (lA = cB): out = out & "cmp.Long2.eq.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: cB = 1.5@: r = (lA < cB): out = out & "cmp.Long2.lt.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: cB = 922337203685477.5807@: r = (lA = cB): out = out & "cmp.Long2.eq.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: lA = 2147483647: cB = 922337203685477.5807@: r = (lA < cB): out = out & "cmp.Long2.lt.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: bB = 0: r = (qA = bB): out = out & "cmp.LongLong0.eq.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: bB = 0: r = (qA < bB): out = out & "cmp.LongLong0.lt.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: bB = 200: r = (qA = bB): out = out & "cmp.LongLong0.eq.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: bB = 200: r = (qA < bB): out = out & "cmp.LongLong0.lt.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: bB = 255: r = (qA = bB): out = out & "cmp.LongLong0.eq.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: bB = 255: r = (qA < bB): out = out & "cmp.LongLong0.lt.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: bB = 0: r = (qA = bB): out = out & "cmp.LongLong1.eq.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: bB = 0: r = (qA < bB): out = out & "cmp.LongLong1.lt.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: bB = 200: r = (qA = bB): out = out & "cmp.LongLong1.eq.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: bB = 200: r = (qA < bB): out = out & "cmp.LongLong1.lt.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: bB = 255: r = (qA = bB): out = out & "cmp.LongLong1.eq.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: bB = 255: r = (qA < bB): out = out & "cmp.LongLong1.lt.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: bB = 0: r = (qA = bB): out = out & "cmp.LongLong2.eq.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: bB = 0: r = (qA < bB): out = out & "cmp.LongLong2.lt.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: bB = 200: r = (qA = bB): out = out & "cmp.LongLong2.eq.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: bB = 200: r = (qA < bB): out = out & "cmp.LongLong2.lt.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: bB = 255: r = (qA = bB): out = out & "cmp.LongLong2.eq.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: bB = 255: r = (qA < bB): out = out & "cmp.LongLong2.lt.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: iB = -32768: r = (qA = iB): out = out & "cmp.LongLong0.eq.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: iB = -32768: r = (qA < iB): out = out & "cmp.LongLong0.lt.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: iB = 1: r = (qA = iB): out = out & "cmp.LongLong0.eq.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: iB = 1: r = (qA < iB): out = out & "cmp.LongLong0.lt.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: iB = 32767: r = (qA = iB): out = out & "cmp.LongLong0.eq.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: iB = 32767: r = (qA < iB): out = out & "cmp.LongLong0.lt.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: iB = -32768: r = (qA = iB): out = out & "cmp.LongLong1.eq.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: iB = -32768: r = (qA < iB): out = out & "cmp.LongLong1.lt.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: iB = 1: r = (qA = iB): out = out & "cmp.LongLong1.eq.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: iB = 1: r = (qA < iB): out = out & "cmp.LongLong1.lt.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: iB = 32767: r = (qA = iB): out = out & "cmp.LongLong1.eq.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: iB = 32767: r = (qA < iB): out = out & "cmp.LongLong1.lt.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: iB = -32768: r = (qA = iB): out = out & "cmp.LongLong2.eq.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: iB = -32768: r = (qA < iB): out = out & "cmp.LongLong2.lt.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: iB = 1: r = (qA = iB): out = out & "cmp.LongLong2.eq.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: iB = 1: r = (qA < iB): out = out & "cmp.LongLong2.lt.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: iB = 32767: r = (qA = iB): out = out & "cmp.LongLong2.eq.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: iB = 32767: r = (qA < iB): out = out & "cmp.LongLong2.lt.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: lB = -2147483648: r = (qA = lB): out = out & "cmp.LongLong0.eq.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: lB = -2147483648: r = (qA < lB): out = out & "cmp.LongLong0.lt.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: lB = 65536: r = (qA = lB): out = out & "cmp.LongLong0.eq.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: lB = 65536: r = (qA < lB): out = out & "cmp.LongLong0.lt.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: lB = 2147483647: r = (qA = lB): out = out & "cmp.LongLong0.eq.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: lB = 2147483647: r = (qA < lB): out = out & "cmp.LongLong0.lt.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: lB = -2147483648: r = (qA = lB): out = out & "cmp.LongLong1.eq.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: lB = -2147483648: r = (qA < lB): out = out & "cmp.LongLong1.lt.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: lB = 65536: r = (qA = lB): out = out & "cmp.LongLong1.eq.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: lB = 65536: r = (qA < lB): out = out & "cmp.LongLong1.lt.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: lB = 2147483647: r = (qA = lB): out = out & "cmp.LongLong1.eq.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: lB = 2147483647: r = (qA < lB): out = out & "cmp.LongLong1.lt.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: lB = -2147483648: r = (qA = lB): out = out & "cmp.LongLong2.eq.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: lB = -2147483648: r = (qA < lB): out = out & "cmp.LongLong2.lt.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: lB = 65536: r = (qA = lB): out = out & "cmp.LongLong2.eq.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: lB = 65536: r = (qA < lB): out = out & "cmp.LongLong2.lt.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: lB = 2147483647: r = (qA = lB): out = out & "cmp.LongLong2.eq.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: lB = 2147483647: r = (qA < lB): out = out & "cmp.LongLong2.lt.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: qB = -9223372036854775807^: r = (qA = qB): out = out & "cmp.LongLong0.eq.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: qB = -9223372036854775807^: r = (qA < qB): out = out & "cmp.LongLong0.lt.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: qB = 3: r = (qA = qB): out = out & "cmp.LongLong0.eq.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: qB = 3: r = (qA < qB): out = out & "cmp.LongLong0.lt.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: qB = 9223372036854775807^: r = (qA = qB): out = out & "cmp.LongLong0.eq.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: qB = 9223372036854775807^: r = (qA < qB): out = out & "cmp.LongLong0.lt.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: qB = -9223372036854775807^: r = (qA = qB): out = out & "cmp.LongLong1.eq.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: qB = -9223372036854775807^: r = (qA < qB): out = out & "cmp.LongLong1.lt.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: qB = 3: r = (qA = qB): out = out & "cmp.LongLong1.eq.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: qB = 3: r = (qA < qB): out = out & "cmp.LongLong1.lt.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: qB = 9223372036854775807^: r = (qA = qB): out = out & "cmp.LongLong1.eq.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: qB = 9223372036854775807^: r = (qA < qB): out = out & "cmp.LongLong1.lt.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: qB = -9223372036854775807^: r = (qA = qB): out = out & "cmp.LongLong2.eq.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: qB = -9223372036854775807^: r = (qA < qB): out = out & "cmp.LongLong2.lt.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: qB = 3: r = (qA = qB): out = out & "cmp.LongLong2.eq.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: qB = 3: r = (qA < qB): out = out & "cmp.LongLong2.lt.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: qB = 9223372036854775807^: r = (qA = qB): out = out & "cmp.LongLong2.eq.LongLong2=" & Rz(Err.Number, r) & vbLf
  P057 = out
End Function
Function P058() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: qA = 9223372036854775807^: qB = 9223372036854775807^: r = (qA < qB): out = out & "cmp.LongLong2.lt.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: sB = 0.5: r = (qA = sB): out = out & "cmp.LongLong0.eq.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: sB = 0.5: r = (qA < sB): out = out & "cmp.LongLong0.lt.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: sB = 16777216: r = (qA = sB): out = out & "cmp.LongLong0.eq.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: sB = 16777216: r = (qA < sB): out = out & "cmp.LongLong0.lt.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: sB = 3.0E+38: r = (qA = sB): out = out & "cmp.LongLong0.eq.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: sB = 3.0E+38: r = (qA < sB): out = out & "cmp.LongLong0.lt.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: sB = 0.5: r = (qA = sB): out = out & "cmp.LongLong1.eq.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: sB = 0.5: r = (qA < sB): out = out & "cmp.LongLong1.lt.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: sB = 16777216: r = (qA = sB): out = out & "cmp.LongLong1.eq.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: sB = 16777216: r = (qA < sB): out = out & "cmp.LongLong1.lt.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: sB = 3.0E+38: r = (qA = sB): out = out & "cmp.LongLong1.eq.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: sB = 3.0E+38: r = (qA < sB): out = out & "cmp.LongLong1.lt.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: sB = 0.5: r = (qA = sB): out = out & "cmp.LongLong2.eq.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: sB = 0.5: r = (qA < sB): out = out & "cmp.LongLong2.lt.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: sB = 16777216: r = (qA = sB): out = out & "cmp.LongLong2.eq.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: sB = 16777216: r = (qA < sB): out = out & "cmp.LongLong2.lt.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: sB = 3.0E+38: r = (qA = sB): out = out & "cmp.LongLong2.eq.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: sB = 3.0E+38: r = (qA < sB): out = out & "cmp.LongLong2.lt.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: dB = 0.5: r = (qA = dB): out = out & "cmp.LongLong0.eq.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: dB = 0.5: r = (qA < dB): out = out & "cmp.LongLong0.lt.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: dB = 2.5: r = (qA = dB): out = out & "cmp.LongLong0.eq.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: dB = 2.5: r = (qA < dB): out = out & "cmp.LongLong0.lt.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: dB = 1.0E+308: r = (qA = dB): out = out & "cmp.LongLong0.eq.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: dB = 1.0E+308: r = (qA < dB): out = out & "cmp.LongLong0.lt.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: dB = 0.5: r = (qA = dB): out = out & "cmp.LongLong1.eq.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: dB = 0.5: r = (qA < dB): out = out & "cmp.LongLong1.lt.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: dB = 2.5: r = (qA = dB): out = out & "cmp.LongLong1.eq.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: dB = 2.5: r = (qA < dB): out = out & "cmp.LongLong1.lt.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: dB = 1.0E+308: r = (qA = dB): out = out & "cmp.LongLong1.eq.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: dB = 1.0E+308: r = (qA < dB): out = out & "cmp.LongLong1.lt.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: dB = 0.5: r = (qA = dB): out = out & "cmp.LongLong2.eq.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: dB = 0.5: r = (qA < dB): out = out & "cmp.LongLong2.lt.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: dB = 2.5: r = (qA = dB): out = out & "cmp.LongLong2.eq.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: dB = 2.5: r = (qA < dB): out = out & "cmp.LongLong2.lt.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: dB = 1.0E+308: r = (qA = dB): out = out & "cmp.LongLong2.eq.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: dB = 1.0E+308: r = (qA < dB): out = out & "cmp.LongLong2.lt.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: cB = -0.5@: r = (qA = cB): out = out & "cmp.LongLong0.eq.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: cB = -0.5@: r = (qA < cB): out = out & "cmp.LongLong0.lt.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: cB = 1.5@: r = (qA = cB): out = out & "cmp.LongLong0.eq.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: cB = 1.5@: r = (qA < cB): out = out & "cmp.LongLong0.lt.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: cB = 922337203685477.5807@: r = (qA = cB): out = out & "cmp.LongLong0.eq.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = -9223372036854775807^: cB = 922337203685477.5807@: r = (qA < cB): out = out & "cmp.LongLong0.lt.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: cB = -0.5@: r = (qA = cB): out = out & "cmp.LongLong1.eq.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: cB = -0.5@: r = (qA < cB): out = out & "cmp.LongLong1.lt.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: cB = 1.5@: r = (qA = cB): out = out & "cmp.LongLong1.eq.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: cB = 1.5@: r = (qA < cB): out = out & "cmp.LongLong1.lt.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: cB = 922337203685477.5807@: r = (qA = cB): out = out & "cmp.LongLong1.eq.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 3: cB = 922337203685477.5807@: r = (qA < cB): out = out & "cmp.LongLong1.lt.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: cB = -0.5@: r = (qA = cB): out = out & "cmp.LongLong2.eq.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: cB = -0.5@: r = (qA < cB): out = out & "cmp.LongLong2.lt.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: cB = 1.5@: r = (qA = cB): out = out & "cmp.LongLong2.eq.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: cB = 1.5@: r = (qA < cB): out = out & "cmp.LongLong2.lt.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: cB = 922337203685477.5807@: r = (qA = cB): out = out & "cmp.LongLong2.eq.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: qA = 9223372036854775807^: cB = 922337203685477.5807@: r = (qA < cB): out = out & "cmp.LongLong2.lt.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: bB = 0: r = (sA = bB): out = out & "cmp.Single0.eq.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: bB = 0: r = (sA < bB): out = out & "cmp.Single0.lt.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: bB = 200: r = (sA = bB): out = out & "cmp.Single0.eq.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: bB = 200: r = (sA < bB): out = out & "cmp.Single0.lt.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: bB = 255: r = (sA = bB): out = out & "cmp.Single0.eq.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: bB = 255: r = (sA < bB): out = out & "cmp.Single0.lt.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: bB = 0: r = (sA = bB): out = out & "cmp.Single1.eq.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: bB = 0: r = (sA < bB): out = out & "cmp.Single1.lt.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: bB = 200: r = (sA = bB): out = out & "cmp.Single1.eq.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: bB = 200: r = (sA < bB): out = out & "cmp.Single1.lt.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: bB = 255: r = (sA = bB): out = out & "cmp.Single1.eq.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: bB = 255: r = (sA < bB): out = out & "cmp.Single1.lt.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: bB = 0: r = (sA = bB): out = out & "cmp.Single2.eq.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: bB = 0: r = (sA < bB): out = out & "cmp.Single2.lt.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: bB = 200: r = (sA = bB): out = out & "cmp.Single2.eq.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: bB = 200: r = (sA < bB): out = out & "cmp.Single2.lt.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: bB = 255: r = (sA = bB): out = out & "cmp.Single2.eq.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: bB = 255: r = (sA < bB): out = out & "cmp.Single2.lt.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: iB = -32768: r = (sA = iB): out = out & "cmp.Single0.eq.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: iB = -32768: r = (sA < iB): out = out & "cmp.Single0.lt.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: iB = 1: r = (sA = iB): out = out & "cmp.Single0.eq.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: iB = 1: r = (sA < iB): out = out & "cmp.Single0.lt.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: iB = 32767: r = (sA = iB): out = out & "cmp.Single0.eq.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: iB = 32767: r = (sA < iB): out = out & "cmp.Single0.lt.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: iB = -32768: r = (sA = iB): out = out & "cmp.Single1.eq.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: iB = -32768: r = (sA < iB): out = out & "cmp.Single1.lt.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: iB = 1: r = (sA = iB): out = out & "cmp.Single1.eq.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: iB = 1: r = (sA < iB): out = out & "cmp.Single1.lt.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: iB = 32767: r = (sA = iB): out = out & "cmp.Single1.eq.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: iB = 32767: r = (sA < iB): out = out & "cmp.Single1.lt.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: iB = -32768: r = (sA = iB): out = out & "cmp.Single2.eq.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: iB = -32768: r = (sA < iB): out = out & "cmp.Single2.lt.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: iB = 1: r = (sA = iB): out = out & "cmp.Single2.eq.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: iB = 1: r = (sA < iB): out = out & "cmp.Single2.lt.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: iB = 32767: r = (sA = iB): out = out & "cmp.Single2.eq.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: iB = 32767: r = (sA < iB): out = out & "cmp.Single2.lt.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: lB = -2147483648: r = (sA = lB): out = out & "cmp.Single0.eq.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: lB = -2147483648: r = (sA < lB): out = out & "cmp.Single0.lt.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: lB = 65536: r = (sA = lB): out = out & "cmp.Single0.eq.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: lB = 65536: r = (sA < lB): out = out & "cmp.Single0.lt.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: lB = 2147483647: r = (sA = lB): out = out & "cmp.Single0.eq.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: lB = 2147483647: r = (sA < lB): out = out & "cmp.Single0.lt.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: lB = -2147483648: r = (sA = lB): out = out & "cmp.Single1.eq.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: lB = -2147483648: r = (sA < lB): out = out & "cmp.Single1.lt.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: lB = 65536: r = (sA = lB): out = out & "cmp.Single1.eq.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: lB = 65536: r = (sA < lB): out = out & "cmp.Single1.lt.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: lB = 2147483647: r = (sA = lB): out = out & "cmp.Single1.eq.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: lB = 2147483647: r = (sA < lB): out = out & "cmp.Single1.lt.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: lB = -2147483648: r = (sA = lB): out = out & "cmp.Single2.eq.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: lB = -2147483648: r = (sA < lB): out = out & "cmp.Single2.lt.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: lB = 65536: r = (sA = lB): out = out & "cmp.Single2.eq.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: lB = 65536: r = (sA < lB): out = out & "cmp.Single2.lt.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: lB = 2147483647: r = (sA = lB): out = out & "cmp.Single2.eq.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: lB = 2147483647: r = (sA < lB): out = out & "cmp.Single2.lt.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: qB = -9223372036854775807^: r = (sA = qB): out = out & "cmp.Single0.eq.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: qB = -9223372036854775807^: r = (sA < qB): out = out & "cmp.Single0.lt.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: qB = 3: r = (sA = qB): out = out & "cmp.Single0.eq.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: qB = 3: r = (sA < qB): out = out & "cmp.Single0.lt.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: qB = 9223372036854775807^: r = (sA = qB): out = out & "cmp.Single0.eq.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: qB = 9223372036854775807^: r = (sA < qB): out = out & "cmp.Single0.lt.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: qB = -9223372036854775807^: r = (sA = qB): out = out & "cmp.Single1.eq.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: qB = -9223372036854775807^: r = (sA < qB): out = out & "cmp.Single1.lt.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: qB = 3: r = (sA = qB): out = out & "cmp.Single1.eq.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: qB = 3: r = (sA < qB): out = out & "cmp.Single1.lt.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: qB = 9223372036854775807^: r = (sA = qB): out = out & "cmp.Single1.eq.LongLong2=" & Rz(Err.Number, r) & vbLf
  P058 = out
End Function
Function P059() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: sA = 16777216: qB = 9223372036854775807^: r = (sA < qB): out = out & "cmp.Single1.lt.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: qB = -9223372036854775807^: r = (sA = qB): out = out & "cmp.Single2.eq.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: qB = -9223372036854775807^: r = (sA < qB): out = out & "cmp.Single2.lt.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: qB = 3: r = (sA = qB): out = out & "cmp.Single2.eq.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: qB = 3: r = (sA < qB): out = out & "cmp.Single2.lt.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: qB = 9223372036854775807^: r = (sA = qB): out = out & "cmp.Single2.eq.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: qB = 9223372036854775807^: r = (sA < qB): out = out & "cmp.Single2.lt.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: sB = 0.5: r = (sA = sB): out = out & "cmp.Single0.eq.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: sB = 0.5: r = (sA < sB): out = out & "cmp.Single0.lt.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: sB = 16777216: r = (sA = sB): out = out & "cmp.Single0.eq.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: sB = 16777216: r = (sA < sB): out = out & "cmp.Single0.lt.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: sB = 3.0E+38: r = (sA = sB): out = out & "cmp.Single0.eq.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: sB = 3.0E+38: r = (sA < sB): out = out & "cmp.Single0.lt.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: sB = 0.5: r = (sA = sB): out = out & "cmp.Single1.eq.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: sB = 0.5: r = (sA < sB): out = out & "cmp.Single1.lt.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: sB = 16777216: r = (sA = sB): out = out & "cmp.Single1.eq.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: sB = 16777216: r = (sA < sB): out = out & "cmp.Single1.lt.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: sB = 3.0E+38: r = (sA = sB): out = out & "cmp.Single1.eq.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: sB = 3.0E+38: r = (sA < sB): out = out & "cmp.Single1.lt.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: sB = 0.5: r = (sA = sB): out = out & "cmp.Single2.eq.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: sB = 0.5: r = (sA < sB): out = out & "cmp.Single2.lt.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: sB = 16777216: r = (sA = sB): out = out & "cmp.Single2.eq.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: sB = 16777216: r = (sA < sB): out = out & "cmp.Single2.lt.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: sB = 3.0E+38: r = (sA = sB): out = out & "cmp.Single2.eq.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: sB = 3.0E+38: r = (sA < sB): out = out & "cmp.Single2.lt.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: dB = 0.5: r = (sA = dB): out = out & "cmp.Single0.eq.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: dB = 0.5: r = (sA < dB): out = out & "cmp.Single0.lt.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: dB = 2.5: r = (sA = dB): out = out & "cmp.Single0.eq.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: dB = 2.5: r = (sA < dB): out = out & "cmp.Single0.lt.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: dB = 1.0E+308: r = (sA = dB): out = out & "cmp.Single0.eq.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: dB = 1.0E+308: r = (sA < dB): out = out & "cmp.Single0.lt.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: dB = 0.5: r = (sA = dB): out = out & "cmp.Single1.eq.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: dB = 0.5: r = (sA < dB): out = out & "cmp.Single1.lt.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: dB = 2.5: r = (sA = dB): out = out & "cmp.Single1.eq.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: dB = 2.5: r = (sA < dB): out = out & "cmp.Single1.lt.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: dB = 1.0E+308: r = (sA = dB): out = out & "cmp.Single1.eq.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: dB = 1.0E+308: r = (sA < dB): out = out & "cmp.Single1.lt.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: dB = 0.5: r = (sA = dB): out = out & "cmp.Single2.eq.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: dB = 0.5: r = (sA < dB): out = out & "cmp.Single2.lt.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: dB = 2.5: r = (sA = dB): out = out & "cmp.Single2.eq.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: dB = 2.5: r = (sA < dB): out = out & "cmp.Single2.lt.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: dB = 1.0E+308: r = (sA = dB): out = out & "cmp.Single2.eq.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: dB = 1.0E+308: r = (sA < dB): out = out & "cmp.Single2.lt.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: cB = -0.5@: r = (sA = cB): out = out & "cmp.Single0.eq.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: cB = -0.5@: r = (sA < cB): out = out & "cmp.Single0.lt.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: cB = 1.5@: r = (sA = cB): out = out & "cmp.Single0.eq.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: cB = 1.5@: r = (sA < cB): out = out & "cmp.Single0.lt.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: cB = 922337203685477.5807@: r = (sA = cB): out = out & "cmp.Single0.eq.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: cB = 922337203685477.5807@: r = (sA < cB): out = out & "cmp.Single0.lt.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: cB = -0.5@: r = (sA = cB): out = out & "cmp.Single1.eq.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: cB = -0.5@: r = (sA < cB): out = out & "cmp.Single1.lt.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: cB = 1.5@: r = (sA = cB): out = out & "cmp.Single1.eq.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: cB = 1.5@: r = (sA < cB): out = out & "cmp.Single1.lt.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: cB = 922337203685477.5807@: r = (sA = cB): out = out & "cmp.Single1.eq.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: cB = 922337203685477.5807@: r = (sA < cB): out = out & "cmp.Single1.lt.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: cB = -0.5@: r = (sA = cB): out = out & "cmp.Single2.eq.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: cB = -0.5@: r = (sA < cB): out = out & "cmp.Single2.lt.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: cB = 1.5@: r = (sA = cB): out = out & "cmp.Single2.eq.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: cB = 1.5@: r = (sA < cB): out = out & "cmp.Single2.lt.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: cB = 922337203685477.5807@: r = (sA = cB): out = out & "cmp.Single2.eq.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: cB = 922337203685477.5807@: r = (sA < cB): out = out & "cmp.Single2.lt.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: bB = 0: r = (dA = bB): out = out & "cmp.Double0.eq.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: bB = 0: r = (dA < bB): out = out & "cmp.Double0.lt.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: bB = 200: r = (dA = bB): out = out & "cmp.Double0.eq.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: bB = 200: r = (dA < bB): out = out & "cmp.Double0.lt.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: bB = 255: r = (dA = bB): out = out & "cmp.Double0.eq.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: bB = 255: r = (dA < bB): out = out & "cmp.Double0.lt.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: bB = 0: r = (dA = bB): out = out & "cmp.Double1.eq.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: bB = 0: r = (dA < bB): out = out & "cmp.Double1.lt.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: bB = 200: r = (dA = bB): out = out & "cmp.Double1.eq.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: bB = 200: r = (dA < bB): out = out & "cmp.Double1.lt.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: bB = 255: r = (dA = bB): out = out & "cmp.Double1.eq.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: bB = 255: r = (dA < bB): out = out & "cmp.Double1.lt.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: bB = 0: r = (dA = bB): out = out & "cmp.Double2.eq.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: bB = 0: r = (dA < bB): out = out & "cmp.Double2.lt.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: bB = 200: r = (dA = bB): out = out & "cmp.Double2.eq.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: bB = 200: r = (dA < bB): out = out & "cmp.Double2.lt.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: bB = 255: r = (dA = bB): out = out & "cmp.Double2.eq.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: bB = 255: r = (dA < bB): out = out & "cmp.Double2.lt.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: iB = -32768: r = (dA = iB): out = out & "cmp.Double0.eq.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: iB = -32768: r = (dA < iB): out = out & "cmp.Double0.lt.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: iB = 1: r = (dA = iB): out = out & "cmp.Double0.eq.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: iB = 1: r = (dA < iB): out = out & "cmp.Double0.lt.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: iB = 32767: r = (dA = iB): out = out & "cmp.Double0.eq.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: iB = 32767: r = (dA < iB): out = out & "cmp.Double0.lt.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: iB = -32768: r = (dA = iB): out = out & "cmp.Double1.eq.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: iB = -32768: r = (dA < iB): out = out & "cmp.Double1.lt.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: iB = 1: r = (dA = iB): out = out & "cmp.Double1.eq.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: iB = 1: r = (dA < iB): out = out & "cmp.Double1.lt.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: iB = 32767: r = (dA = iB): out = out & "cmp.Double1.eq.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: iB = 32767: r = (dA < iB): out = out & "cmp.Double1.lt.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: iB = -32768: r = (dA = iB): out = out & "cmp.Double2.eq.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: iB = -32768: r = (dA < iB): out = out & "cmp.Double2.lt.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: iB = 1: r = (dA = iB): out = out & "cmp.Double2.eq.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: iB = 1: r = (dA < iB): out = out & "cmp.Double2.lt.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: iB = 32767: r = (dA = iB): out = out & "cmp.Double2.eq.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: iB = 32767: r = (dA < iB): out = out & "cmp.Double2.lt.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: lB = -2147483648: r = (dA = lB): out = out & "cmp.Double0.eq.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: lB = -2147483648: r = (dA < lB): out = out & "cmp.Double0.lt.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: lB = 65536: r = (dA = lB): out = out & "cmp.Double0.eq.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: lB = 65536: r = (dA < lB): out = out & "cmp.Double0.lt.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: lB = 2147483647: r = (dA = lB): out = out & "cmp.Double0.eq.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: lB = 2147483647: r = (dA < lB): out = out & "cmp.Double0.lt.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: lB = -2147483648: r = (dA = lB): out = out & "cmp.Double1.eq.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: lB = -2147483648: r = (dA < lB): out = out & "cmp.Double1.lt.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: lB = 65536: r = (dA = lB): out = out & "cmp.Double1.eq.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: lB = 65536: r = (dA < lB): out = out & "cmp.Double1.lt.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: lB = 2147483647: r = (dA = lB): out = out & "cmp.Double1.eq.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: lB = 2147483647: r = (dA < lB): out = out & "cmp.Double1.lt.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: lB = -2147483648: r = (dA = lB): out = out & "cmp.Double2.eq.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: lB = -2147483648: r = (dA < lB): out = out & "cmp.Double2.lt.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: lB = 65536: r = (dA = lB): out = out & "cmp.Double2.eq.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: lB = 65536: r = (dA < lB): out = out & "cmp.Double2.lt.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: lB = 2147483647: r = (dA = lB): out = out & "cmp.Double2.eq.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: lB = 2147483647: r = (dA < lB): out = out & "cmp.Double2.lt.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: qB = -9223372036854775807^: r = (dA = qB): out = out & "cmp.Double0.eq.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: qB = -9223372036854775807^: r = (dA < qB): out = out & "cmp.Double0.lt.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: qB = 3: r = (dA = qB): out = out & "cmp.Double0.eq.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: qB = 3: r = (dA < qB): out = out & "cmp.Double0.lt.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: qB = 9223372036854775807^: r = (dA = qB): out = out & "cmp.Double0.eq.LongLong2=" & Rz(Err.Number, r) & vbLf
  P059 = out
End Function
Function P060() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: dA = 0.5: qB = 9223372036854775807^: r = (dA < qB): out = out & "cmp.Double0.lt.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: qB = -9223372036854775807^: r = (dA = qB): out = out & "cmp.Double1.eq.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: qB = -9223372036854775807^: r = (dA < qB): out = out & "cmp.Double1.lt.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: qB = 3: r = (dA = qB): out = out & "cmp.Double1.eq.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: qB = 3: r = (dA < qB): out = out & "cmp.Double1.lt.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: qB = 9223372036854775807^: r = (dA = qB): out = out & "cmp.Double1.eq.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: qB = 9223372036854775807^: r = (dA < qB): out = out & "cmp.Double1.lt.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: qB = -9223372036854775807^: r = (dA = qB): out = out & "cmp.Double2.eq.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: qB = -9223372036854775807^: r = (dA < qB): out = out & "cmp.Double2.lt.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: qB = 3: r = (dA = qB): out = out & "cmp.Double2.eq.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: qB = 3: r = (dA < qB): out = out & "cmp.Double2.lt.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: qB = 9223372036854775807^: r = (dA = qB): out = out & "cmp.Double2.eq.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: qB = 9223372036854775807^: r = (dA < qB): out = out & "cmp.Double2.lt.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: sB = 0.5: r = (dA = sB): out = out & "cmp.Double0.eq.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: sB = 0.5: r = (dA < sB): out = out & "cmp.Double0.lt.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: sB = 16777216: r = (dA = sB): out = out & "cmp.Double0.eq.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: sB = 16777216: r = (dA < sB): out = out & "cmp.Double0.lt.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: sB = 3.0E+38: r = (dA = sB): out = out & "cmp.Double0.eq.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: sB = 3.0E+38: r = (dA < sB): out = out & "cmp.Double0.lt.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: sB = 0.5: r = (dA = sB): out = out & "cmp.Double1.eq.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: sB = 0.5: r = (dA < sB): out = out & "cmp.Double1.lt.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: sB = 16777216: r = (dA = sB): out = out & "cmp.Double1.eq.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: sB = 16777216: r = (dA < sB): out = out & "cmp.Double1.lt.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: sB = 3.0E+38: r = (dA = sB): out = out & "cmp.Double1.eq.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: sB = 3.0E+38: r = (dA < sB): out = out & "cmp.Double1.lt.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: sB = 0.5: r = (dA = sB): out = out & "cmp.Double2.eq.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: sB = 0.5: r = (dA < sB): out = out & "cmp.Double2.lt.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: sB = 16777216: r = (dA = sB): out = out & "cmp.Double2.eq.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: sB = 16777216: r = (dA < sB): out = out & "cmp.Double2.lt.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: sB = 3.0E+38: r = (dA = sB): out = out & "cmp.Double2.eq.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: sB = 3.0E+38: r = (dA < sB): out = out & "cmp.Double2.lt.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: dB = 0.5: r = (dA = dB): out = out & "cmp.Double0.eq.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: dB = 0.5: r = (dA < dB): out = out & "cmp.Double0.lt.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: dB = 2.5: r = (dA = dB): out = out & "cmp.Double0.eq.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: dB = 2.5: r = (dA < dB): out = out & "cmp.Double0.lt.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: dB = 1.0E+308: r = (dA = dB): out = out & "cmp.Double0.eq.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: dB = 1.0E+308: r = (dA < dB): out = out & "cmp.Double0.lt.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: dB = 0.5: r = (dA = dB): out = out & "cmp.Double1.eq.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: dB = 0.5: r = (dA < dB): out = out & "cmp.Double1.lt.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: dB = 2.5: r = (dA = dB): out = out & "cmp.Double1.eq.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: dB = 2.5: r = (dA < dB): out = out & "cmp.Double1.lt.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: dB = 1.0E+308: r = (dA = dB): out = out & "cmp.Double1.eq.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: dB = 1.0E+308: r = (dA < dB): out = out & "cmp.Double1.lt.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: dB = 0.5: r = (dA = dB): out = out & "cmp.Double2.eq.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: dB = 0.5: r = (dA < dB): out = out & "cmp.Double2.lt.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: dB = 2.5: r = (dA = dB): out = out & "cmp.Double2.eq.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: dB = 2.5: r = (dA < dB): out = out & "cmp.Double2.lt.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: dB = 1.0E+308: r = (dA = dB): out = out & "cmp.Double2.eq.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: dB = 1.0E+308: r = (dA < dB): out = out & "cmp.Double2.lt.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: cB = -0.5@: r = (dA = cB): out = out & "cmp.Double0.eq.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: cB = -0.5@: r = (dA < cB): out = out & "cmp.Double0.lt.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: cB = 1.5@: r = (dA = cB): out = out & "cmp.Double0.eq.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: cB = 1.5@: r = (dA < cB): out = out & "cmp.Double0.lt.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: cB = 922337203685477.5807@: r = (dA = cB): out = out & "cmp.Double0.eq.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: cB = 922337203685477.5807@: r = (dA < cB): out = out & "cmp.Double0.lt.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: cB = -0.5@: r = (dA = cB): out = out & "cmp.Double1.eq.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: cB = -0.5@: r = (dA < cB): out = out & "cmp.Double1.lt.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: cB = 1.5@: r = (dA = cB): out = out & "cmp.Double1.eq.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: cB = 1.5@: r = (dA < cB): out = out & "cmp.Double1.lt.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: cB = 922337203685477.5807@: r = (dA = cB): out = out & "cmp.Double1.eq.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: cB = 922337203685477.5807@: r = (dA < cB): out = out & "cmp.Double1.lt.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: cB = -0.5@: r = (dA = cB): out = out & "cmp.Double2.eq.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: cB = -0.5@: r = (dA < cB): out = out & "cmp.Double2.lt.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: cB = 1.5@: r = (dA = cB): out = out & "cmp.Double2.eq.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: cB = 1.5@: r = (dA < cB): out = out & "cmp.Double2.lt.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: cB = 922337203685477.5807@: r = (dA = cB): out = out & "cmp.Double2.eq.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: cB = 922337203685477.5807@: r = (dA < cB): out = out & "cmp.Double2.lt.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: bB = 0: r = (cA = bB): out = out & "cmp.Currency0.eq.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: bB = 0: r = (cA < bB): out = out & "cmp.Currency0.lt.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: bB = 200: r = (cA = bB): out = out & "cmp.Currency0.eq.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: bB = 200: r = (cA < bB): out = out & "cmp.Currency0.lt.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: bB = 255: r = (cA = bB): out = out & "cmp.Currency0.eq.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: bB = 255: r = (cA < bB): out = out & "cmp.Currency0.lt.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: bB = 0: r = (cA = bB): out = out & "cmp.Currency1.eq.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: bB = 0: r = (cA < bB): out = out & "cmp.Currency1.lt.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: bB = 200: r = (cA = bB): out = out & "cmp.Currency1.eq.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: bB = 200: r = (cA < bB): out = out & "cmp.Currency1.lt.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: bB = 255: r = (cA = bB): out = out & "cmp.Currency1.eq.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: bB = 255: r = (cA < bB): out = out & "cmp.Currency1.lt.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bB = 0: r = (cA = bB): out = out & "cmp.Currency2.eq.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bB = 0: r = (cA < bB): out = out & "cmp.Currency2.lt.Byte0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bB = 200: r = (cA = bB): out = out & "cmp.Currency2.eq.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bB = 200: r = (cA < bB): out = out & "cmp.Currency2.lt.Byte1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bB = 255: r = (cA = bB): out = out & "cmp.Currency2.eq.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: bB = 255: r = (cA < bB): out = out & "cmp.Currency2.lt.Byte2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: iB = -32768: r = (cA = iB): out = out & "cmp.Currency0.eq.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: iB = -32768: r = (cA < iB): out = out & "cmp.Currency0.lt.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: iB = 1: r = (cA = iB): out = out & "cmp.Currency0.eq.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: iB = 1: r = (cA < iB): out = out & "cmp.Currency0.lt.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: iB = 32767: r = (cA = iB): out = out & "cmp.Currency0.eq.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: iB = 32767: r = (cA < iB): out = out & "cmp.Currency0.lt.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: iB = -32768: r = (cA = iB): out = out & "cmp.Currency1.eq.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: iB = -32768: r = (cA < iB): out = out & "cmp.Currency1.lt.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: iB = 1: r = (cA = iB): out = out & "cmp.Currency1.eq.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: iB = 1: r = (cA < iB): out = out & "cmp.Currency1.lt.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: iB = 32767: r = (cA = iB): out = out & "cmp.Currency1.eq.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: iB = 32767: r = (cA < iB): out = out & "cmp.Currency1.lt.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iB = -32768: r = (cA = iB): out = out & "cmp.Currency2.eq.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iB = -32768: r = (cA < iB): out = out & "cmp.Currency2.lt.Integer0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iB = 1: r = (cA = iB): out = out & "cmp.Currency2.eq.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iB = 1: r = (cA < iB): out = out & "cmp.Currency2.lt.Integer1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iB = 32767: r = (cA = iB): out = out & "cmp.Currency2.eq.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: iB = 32767: r = (cA < iB): out = out & "cmp.Currency2.lt.Integer2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: lB = -2147483648: r = (cA = lB): out = out & "cmp.Currency0.eq.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: lB = -2147483648: r = (cA < lB): out = out & "cmp.Currency0.lt.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: lB = 65536: r = (cA = lB): out = out & "cmp.Currency0.eq.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: lB = 65536: r = (cA < lB): out = out & "cmp.Currency0.lt.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: lB = 2147483647: r = (cA = lB): out = out & "cmp.Currency0.eq.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: lB = 2147483647: r = (cA < lB): out = out & "cmp.Currency0.lt.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: lB = -2147483648: r = (cA = lB): out = out & "cmp.Currency1.eq.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: lB = -2147483648: r = (cA < lB): out = out & "cmp.Currency1.lt.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: lB = 65536: r = (cA = lB): out = out & "cmp.Currency1.eq.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: lB = 65536: r = (cA < lB): out = out & "cmp.Currency1.lt.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: lB = 2147483647: r = (cA = lB): out = out & "cmp.Currency1.eq.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: lB = 2147483647: r = (cA < lB): out = out & "cmp.Currency1.lt.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lB = -2147483648: r = (cA = lB): out = out & "cmp.Currency2.eq.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lB = -2147483648: r = (cA < lB): out = out & "cmp.Currency2.lt.Long0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lB = 65536: r = (cA = lB): out = out & "cmp.Currency2.eq.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lB = 65536: r = (cA < lB): out = out & "cmp.Currency2.lt.Long1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: lB = 2147483647: r = (cA = lB): out = out & "cmp.Currency2.eq.Long2=" & Rz(Err.Number, r) & vbLf
  P060 = out
End Function
Function P061() As String
  Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String
  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency
  On Error Resume Next
  Err.Clear: cA = 922337203685477.5807@: lB = 2147483647: r = (cA < lB): out = out & "cmp.Currency2.lt.Long2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: qB = -9223372036854775807^: r = (cA = qB): out = out & "cmp.Currency0.eq.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: qB = -9223372036854775807^: r = (cA < qB): out = out & "cmp.Currency0.lt.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: qB = 3: r = (cA = qB): out = out & "cmp.Currency0.eq.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: qB = 3: r = (cA < qB): out = out & "cmp.Currency0.lt.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: qB = 9223372036854775807^: r = (cA = qB): out = out & "cmp.Currency0.eq.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: qB = 9223372036854775807^: r = (cA < qB): out = out & "cmp.Currency0.lt.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: qB = -9223372036854775807^: r = (cA = qB): out = out & "cmp.Currency1.eq.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: qB = -9223372036854775807^: r = (cA < qB): out = out & "cmp.Currency1.lt.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: qB = 3: r = (cA = qB): out = out & "cmp.Currency1.eq.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: qB = 3: r = (cA < qB): out = out & "cmp.Currency1.lt.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: qB = 9223372036854775807^: r = (cA = qB): out = out & "cmp.Currency1.eq.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: qB = 9223372036854775807^: r = (cA < qB): out = out & "cmp.Currency1.lt.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qB = -9223372036854775807^: r = (cA = qB): out = out & "cmp.Currency2.eq.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qB = -9223372036854775807^: r = (cA < qB): out = out & "cmp.Currency2.lt.LongLong0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qB = 3: r = (cA = qB): out = out & "cmp.Currency2.eq.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qB = 3: r = (cA < qB): out = out & "cmp.Currency2.lt.LongLong1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qB = 9223372036854775807^: r = (cA = qB): out = out & "cmp.Currency2.eq.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: qB = 9223372036854775807^: r = (cA < qB): out = out & "cmp.Currency2.lt.LongLong2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: sB = 0.5: r = (cA = sB): out = out & "cmp.Currency0.eq.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: sB = 0.5: r = (cA < sB): out = out & "cmp.Currency0.lt.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: sB = 16777216: r = (cA = sB): out = out & "cmp.Currency0.eq.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: sB = 16777216: r = (cA < sB): out = out & "cmp.Currency0.lt.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: sB = 3.0E+38: r = (cA = sB): out = out & "cmp.Currency0.eq.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: sB = 3.0E+38: r = (cA < sB): out = out & "cmp.Currency0.lt.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: sB = 0.5: r = (cA = sB): out = out & "cmp.Currency1.eq.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: sB = 0.5: r = (cA < sB): out = out & "cmp.Currency1.lt.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: sB = 16777216: r = (cA = sB): out = out & "cmp.Currency1.eq.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: sB = 16777216: r = (cA < sB): out = out & "cmp.Currency1.lt.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: sB = 3.0E+38: r = (cA = sB): out = out & "cmp.Currency1.eq.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: sB = 3.0E+38: r = (cA < sB): out = out & "cmp.Currency1.lt.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sB = 0.5: r = (cA = sB): out = out & "cmp.Currency2.eq.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sB = 0.5: r = (cA < sB): out = out & "cmp.Currency2.lt.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sB = 16777216: r = (cA = sB): out = out & "cmp.Currency2.eq.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sB = 16777216: r = (cA < sB): out = out & "cmp.Currency2.lt.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sB = 3.0E+38: r = (cA = sB): out = out & "cmp.Currency2.eq.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: sB = 3.0E+38: r = (cA < sB): out = out & "cmp.Currency2.lt.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: dB = 0.5: r = (cA = dB): out = out & "cmp.Currency0.eq.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: dB = 0.5: r = (cA < dB): out = out & "cmp.Currency0.lt.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: dB = 2.5: r = (cA = dB): out = out & "cmp.Currency0.eq.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: dB = 2.5: r = (cA < dB): out = out & "cmp.Currency0.lt.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: dB = 1.0E+308: r = (cA = dB): out = out & "cmp.Currency0.eq.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: dB = 1.0E+308: r = (cA < dB): out = out & "cmp.Currency0.lt.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: dB = 0.5: r = (cA = dB): out = out & "cmp.Currency1.eq.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: dB = 0.5: r = (cA < dB): out = out & "cmp.Currency1.lt.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: dB = 2.5: r = (cA = dB): out = out & "cmp.Currency1.eq.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: dB = 2.5: r = (cA < dB): out = out & "cmp.Currency1.lt.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: dB = 1.0E+308: r = (cA = dB): out = out & "cmp.Currency1.eq.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: dB = 1.0E+308: r = (cA < dB): out = out & "cmp.Currency1.lt.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dB = 0.5: r = (cA = dB): out = out & "cmp.Currency2.eq.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dB = 0.5: r = (cA < dB): out = out & "cmp.Currency2.lt.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dB = 2.5: r = (cA = dB): out = out & "cmp.Currency2.eq.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dB = 2.5: r = (cA < dB): out = out & "cmp.Currency2.lt.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dB = 1.0E+308: r = (cA = dB): out = out & "cmp.Currency2.eq.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: dB = 1.0E+308: r = (cA < dB): out = out & "cmp.Currency2.lt.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: cB = -0.5@: r = (cA = cB): out = out & "cmp.Currency0.eq.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: cB = -0.5@: r = (cA < cB): out = out & "cmp.Currency0.lt.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: cB = 1.5@: r = (cA = cB): out = out & "cmp.Currency0.eq.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: cB = 1.5@: r = (cA < cB): out = out & "cmp.Currency0.lt.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: cB = 922337203685477.5807@: r = (cA = cB): out = out & "cmp.Currency0.eq.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: cB = 922337203685477.5807@: r = (cA < cB): out = out & "cmp.Currency0.lt.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: cB = -0.5@: r = (cA = cB): out = out & "cmp.Currency1.eq.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: cB = -0.5@: r = (cA < cB): out = out & "cmp.Currency1.lt.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: cB = 1.5@: r = (cA = cB): out = out & "cmp.Currency1.eq.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: cB = 1.5@: r = (cA < cB): out = out & "cmp.Currency1.lt.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: cB = 922337203685477.5807@: r = (cA = cB): out = out & "cmp.Currency1.eq.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: cB = 922337203685477.5807@: r = (cA < cB): out = out & "cmp.Currency1.lt.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cB = -0.5@: r = (cA = cB): out = out & "cmp.Currency2.eq.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cB = -0.5@: r = (cA < cB): out = out & "cmp.Currency2.lt.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cB = 1.5@: r = (cA = cB): out = out & "cmp.Currency2.eq.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cB = 1.5@: r = (cA < cB): out = out & "cmp.Currency2.lt.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cB = 922337203685477.5807@: r = (cA = cB): out = out & "cmp.Currency2.eq.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: cB = 922337203685477.5807@: r = (cA < cB): out = out & "cmp.Currency2.lt.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: r = Int(sA): out = out & "fn.Int.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: r = Int(sA): out = out & "fn.Int.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: r = Int(sA): out = out & "fn.Int.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: r = Int(dA): out = out & "fn.Int.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: r = Int(dA): out = out & "fn.Int.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: r = Int(dA): out = out & "fn.Int.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: r = Int(cA): out = out & "fn.Int.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: r = Int(cA): out = out & "fn.Int.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: r = Int(cA): out = out & "fn.Int.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: r = Fix(sA): out = out & "fn.Fix.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: r = Fix(sA): out = out & "fn.Fix.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: r = Fix(sA): out = out & "fn.Fix.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: r = Fix(dA): out = out & "fn.Fix.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: r = Fix(dA): out = out & "fn.Fix.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: r = Fix(dA): out = out & "fn.Fix.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: r = Fix(cA): out = out & "fn.Fix.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: r = Fix(cA): out = out & "fn.Fix.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: r = Fix(cA): out = out & "fn.Fix.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: r = Sgn(sA): out = out & "fn.Sgn.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: r = Sgn(sA): out = out & "fn.Sgn.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: r = Sgn(sA): out = out & "fn.Sgn.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: r = Sgn(dA): out = out & "fn.Sgn.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: r = Sgn(dA): out = out & "fn.Sgn.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: r = Sgn(dA): out = out & "fn.Sgn.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: r = Sgn(cA): out = out & "fn.Sgn.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: r = Sgn(cA): out = out & "fn.Sgn.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: r = Sgn(cA): out = out & "fn.Sgn.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: r = CInt(sA): out = out & "fn.CInt.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: r = CInt(sA): out = out & "fn.CInt.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: r = CInt(sA): out = out & "fn.CInt.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: r = CInt(dA): out = out & "fn.CInt.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: r = CInt(dA): out = out & "fn.CInt.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: r = CInt(dA): out = out & "fn.CInt.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: r = CInt(cA): out = out & "fn.CInt.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: r = CInt(cA): out = out & "fn.CInt.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: r = CInt(cA): out = out & "fn.CInt.Currency2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 0.5: r = Round(sA): out = out & "fn.Round.Single0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 16777216: r = Round(sA): out = out & "fn.Round.Single1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: sA = 3.0E+38: r = Round(sA): out = out & "fn.Round.Single2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 0.5: r = Round(dA): out = out & "fn.Round.Double0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 2.5: r = Round(dA): out = out & "fn.Round.Double1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: dA = 1.0E+308: r = Round(dA): out = out & "fn.Round.Double2=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = -0.5@: r = Round(cA): out = out & "fn.Round.Currency0=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 1.5@: r = Round(cA): out = out & "fn.Round.Currency1=" & Rz(Err.Number, r) & vbLf
  Err.Clear: cA = 922337203685477.5807@: r = Round(cA): out = out & "fn.Round.Currency2=" & Rz(Err.Number, r) & vbLf
  P061 = out
End Function
Public Function RunAll() As String
  Dim s As String
  s = s & P000()
  s = s & P001()
  s = s & P002()
  s = s & P003()
  s = s & P004()
  s = s & P005()
  s = s & P006()
  s = s & P007()
  s = s & P008()
  s = s & P009()
  s = s & P010()
  s = s & P011()
  s = s & P012()
  s = s & P013()
  s = s & P014()
  s = s & P015()
  s = s & P016()
  s = s & P017()
  s = s & P018()
  s = s & P019()
  s = s & P020()
  s = s & P021()
  s = s & P022()
  s = s & P023()
  s = s & P024()
  s = s & P025()
  s = s & P026()
  s = s & P027()
  s = s & P028()
  s = s & P029()
  s = s & P030()
  s = s & P031()
  s = s & P032()
  s = s & P033()
  s = s & P034()
  s = s & P035()
  s = s & P036()
  s = s & P037()
  s = s & P038()
  s = s & P039()
  s = s & P040()
  s = s & P041()
  s = s & P042()
  s = s & P043()
  s = s & P044()
  s = s & P045()
  s = s & P046()
  s = s & P047()
  s = s & P048()
  s = s & P049()
  s = s & P050()
  s = s & P051()
  s = s & P052()
  s = s & P053()
  s = s & P054()
  s = s & P055()
  s = s & P056()
  s = s & P057()
  s = s & P058()
  s = s & P059()
  s = s & P060()
  s = s & P061()
  RunAll = s
End Function
