#!/usr/bin/env python3
"""Generates double-string.vba: Excel VBA probe for Double and Single to String conversion."""
import sys
D = ["0", "-1", "1", "0.5", "-0.5", "0.1", "0.2", "0.1 + 0.2", "1 / 3", "2 / 3", "10 / 3", "100", "1234.5", "123456.789",
     "123456789012345", "123456789012345.6", "1234567890123456", "12345678901234567", "99999999999999.9", "999999999999999", "9999999999999999",
     "100000000000000", "1E+15", "1E+16", "1.5E+15", "1E+22", "1E+100", "1.7976931348623157E+308", "1E-4", "1.5E-4", "1E-5", "1.5E-5",
     "0.00012345", "0.000012345", "1E-10", "1.5E-10", "1E-300", "4.94065645841247E-324", "2 ^ 53", "2 ^ 63", "2 ^ 31", "-2 ^ 31", "0.1 * 3",
     "1 - 0.9", "3.14159265358979", "2.718281828459045", "1.0000000000000002", "0.99999999999999989", "0.30000000000000004", "123456789.123456789",
     "5E-5", "9.5E-5", "99999.99999999999", "0.000099999999999999", "4503599627370496.5", "1E+15 + 0.5", "-1E+15", "-1.5E-5", "1E+14 + 0.5", "255.5", "32767.5", "2147483647.5"]
S = ["0.5", "0.1", "1 / 3", "2 / 3", "16777216", "16777217", "1234567", "12345678", "123456789", "1E+7", "1E+8", "1E-4", "1E-5", "1.5E-5", "3E+38", "1E-38", "1.401298E-45",
     "0.1 + 0.2", "123456.7", "1234567.8", "99999.99", "0.0001234", "0.00001234", "100000", "1000000", "10000000", "-0.5", "-1E+7", "16777216 + 4", "2147483648"]
out = []
def proc(name, items, dec, conv):
    out.append(f"Function {name}() As String\n  Dim x As {dec}, s As String, t As String, u As String\n  On Error Resume Next\n")
    for i, e in enumerate(items):
        out.append(f'  Err.Clear: x = {e}: s = CStr(x): t = Str$(x): u = "" & x\n  {name} = {name} & "{name}.{i}=" & s & "|" & t & "|" & u & "|" & Err.Number & vbLf\n')
    out.append("End Function\n")
proc("DblStr", D, "Double", "")
proc("SglStr", S, "Single", "")
# negative zero and special double cases
out.append('Function ZeroStr() As String\n  Dim x As Double, y As Double\n  On Error Resume Next\n  x = 0: y = -x\n  ZeroStr = "negzero=" & CStr(y) & "|" & Str$(y) & "|" & ("" & y) & "|" & (1 / y) & "|" & Err.Number & vbLf\n  Err.Clear: y = -1: y = y * 0: ZeroStr = ZeroStr & "mulzero=" & CStr(y) & "|" & Str$(y) & "|" & Err.Number & vbLf\nEnd Function\n')
out.append('Function RunAll() As String\n  RunAll = DblStr() & SglStr() & ZeroStr()\nEnd Function\n')
sys.stdout.write("".join(out).replace("\n", "\r\n"))
