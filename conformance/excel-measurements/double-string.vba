Function DblStr() As String
  Dim x As Double, s As String, t As String, u As String
  On Error Resume Next
  Err.Clear: x = 0: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.0=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = -1: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.1=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.2=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 0.5: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.3=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = -0.5: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.4=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 0.1: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.5=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 0.2: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.6=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 0.1 + 0.2: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.7=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1 / 3: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.8=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 2 / 3: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.9=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 10 / 3: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.10=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 100: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.11=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1234.5: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.12=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 123456.789: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.13=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 123456789012345: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.14=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 123456789012345.6: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.15=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1234567890123456: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.16=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 12345678901234567: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.17=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 99999999999999.9: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.18=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 999999999999999: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.19=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 9999999999999999: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.20=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 100000000000000: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.21=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1E+15: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.22=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1E+16: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.23=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E+15: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.24=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1E+22: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.25=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1E+100: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.26=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1.7976931348623157E+308: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.27=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-4: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.28=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-4: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.29=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-5: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.30=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-5: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.31=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 0.00012345: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.32=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 0.000012345: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.33=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-10: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.34=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-10: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.35=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-300: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.36=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 4.94065645841247E-324: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.37=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 2 ^ 53: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.38=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 2 ^ 63: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.39=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 2 ^ 31: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.40=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = -2 ^ 31: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.41=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 0.1 * 3: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.42=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1 - 0.9: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.43=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 3.14159265358979: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.44=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 2.718281828459045: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.45=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1.0000000000000002: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.46=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 0.99999999999999989: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.47=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 0.30000000000000004: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.48=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 123456789.123456789: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.49=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 5E-5: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.50=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 9.5E-5: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.51=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 99999.99999999999: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.52=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 0.000099999999999999: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.53=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 4503599627370496.5: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.54=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1E+15 + 0.5: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.55=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = -1E+15: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.56=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = -1.5E-5: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.57=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1E+14 + 0.5: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.58=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 255.5: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.59=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 32767.5: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.60=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 2147483647.5: s = CStr(x): t = Str$(x): u = "" & x
  DblStr = DblStr & "DblStr.61=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
End Function
Function SglStr() As String
  Dim x As Single, s As String, t As String, u As String
  On Error Resume Next
  Err.Clear: x = 0.5: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.0=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 0.1: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.1=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1 / 3: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.2=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 2 / 3: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.3=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 16777216: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.4=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 16777217: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.5=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1234567: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.6=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 12345678: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.7=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 123456789: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.8=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1E+7: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.9=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1E+8: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.10=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-4: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.11=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-5: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.12=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-5: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.13=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 3E+38: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.14=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-38: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.15=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1.401298E-45: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.16=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 0.1 + 0.2: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.17=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 123456.7: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.18=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1234567.8: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.19=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 99999.99: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.20=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 0.0001234: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.21=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 0.00001234: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.22=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 100000: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.23=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 1000000: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.24=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 10000000: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.25=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = -0.5: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.26=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = -1E+7: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.27=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 16777216 + 4: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.28=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
  Err.Clear: x = 2147483648: s = CStr(x): t = Str$(x): u = "" & x
  SglStr = SglStr & "SglStr.29=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf
End Function
Function ZeroStr() As String
  Dim x As Double, y As Double
  On Error Resume Next
  x = 0: y = -x
  ZeroStr = "negzero=" & CStr(y) & "|" & Str$(y) & "|" & ("" & y) & "|" & (1 / y) & "|" & Err.Number & vbLf
  Err.Clear: y = -1: y = y * 0: ZeroStr = ZeroStr & "mulzero=" & CStr(y) & "|" & Str$(y) & "|" & Err.Number & vbLf
End Function
Function RunAll() As String
  RunAll = DblStr() & SglStr() & ZeroStr()
End Function
