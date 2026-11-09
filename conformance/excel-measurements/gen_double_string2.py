#!/usr/bin/env python3
"""Generates double-string-bounds.vba: where Double and Single string conversion switches to E notation."""
import sys
out = []
def proc(name, dec, items):
    out.append(f"Function {name}() As String\n  Dim x As {dec}, s As String, t As String\n  On Error Resume Next\n")
    for i, e in enumerate(items):
        out.append(f'  Err.Clear: x = {e}: s = CStr(x): t = Str$(x)\n  {name} = {name} & "{name}.{e.replace(" ", "")}=" & s & "|" & t & "|" & Err.Number & vbLf\n')
    out.append("End Function\n")
D = []
for k in range(3, 24):
    for m in ("1", "1.5", "1.2345", "1.23456789012345"):
        D.append(f"{m}E-{k:02d}")
for k in range(13, 20):
    for m in ("1", "1.5", "1.2345", "1.23456789012345"):
        D.append(f"{m}E+{k:02d}")
S = []
for k in range(3, 14):
    for m in ("1", "1.5", "1.2345", "1.234567"):
        S.append(f"{m}E-{k:02d}")
for k in range(4, 12):
    for m in ("1", "1.5", "1.2345", "1.234567"):
        S.append(f"{m}E+{k:02d}")
proc("DblB", "Double", D)
proc("SglB", "Single", S)
out.append('Function ZeroB() As String\n  Dim x As Double, y As Double, z As Double\n  On Error Resume Next\n  x = 0: y = -1: z = x * y\n  ZeroB = "negzero=" & CStr(z) & "|" & Str$(z) & "|" & ("" & z) & "|" & Err.Number & vbLf\n  Err.Clear: z = -x\n  ZeroB = ZeroB & "negzero2=" & CStr(z) & "|" & Str$(z) & "|" & Err.Number & vbLf\n  Err.Clear: z = CDbl(-0#)\n  ZeroB = ZeroB & "negzero3=" & CStr(z) & "|" & Str$(z) & "|" & Err.Number & vbLf\nEnd Function\n')
out.append('Function RunAll() As String\n  RunAll = DblB() & SglB() & ZeroB()\nEnd Function\n')
sys.stdout.write("".join(out).replace("\n", "\r\n"))
