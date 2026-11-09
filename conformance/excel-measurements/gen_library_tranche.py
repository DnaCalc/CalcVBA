#!/usr/bin/env python3
"""Generates library-tranche.vba: early library oracle tranche (string, math, conversion functions).
Each case runs under On Error Resume Next and records id=TypeName:value or id=ERR<n>."""
import sys
exprs = []
def add(id_, e):
    exprs.append((id_, e))
S = ['""', '"a"', '"abc"', '"Hello World"', '"  pad  "', '"12"', '"-3.5"', '"1e3"', '"&H1F"', '"abc abc"', 'Null', 'Empty']
for i, s in enumerate(S):
    add(f"Len.{i}", f"Len(vA{i})")
for i, s in enumerate(S):
    for fn in ("UCase", "LCase", "Trim", "LTrim", "RTrim"):
        add(f"{fn}.{i}", f"{fn}(vA{i})")
    add(f"Asc.{i}", f"Asc(vA{i})")
    add(f"StrReverse.{i}", f"StrReverse(vA{i})")
for i in range(len(S)):
    for n in (-1, 0, 1, 2, 5, 100):
        add(f"Left.{i}.{n}", f"Left(vA{i}, {n})")
        add(f"Right.{i}.{n}", f"Right(vA{i}, {n})")
        add(f"Mid.{i}.{n}", f"Mid(vA{i}, {n})")
        add(f"Mid3.{i}.{n}", f"Mid(vA{i}, {n}, 2)")
for i in (2, 3, 7, 9):
    for j in (0, 1, 2, 3, 9):
        for st in (-1, 0, 1, 3, 20):
            add(f"InStr.{i}.{j}.{st}", f"InStr({st}, vA{i}, vA{j})")
        add(f"InStrRev.{i}.{j}", f"InStrRev(vA{i}, vA{j})")
        add(f"StrComp.{i}.{j}", f"StrComp(vA{i}, vA{j})")
        add(f"StrCompT.{i}.{j}", f"StrComp(vA{i}, vA{j}, vbTextCompare)")
        add(f"Replace.{i}.{j}", f'Replace(vA{i}, vA{j}, "X")')
for n in (-1, 0, 1, 3):
    add(f"String.{n}", f'String({n}, "ab")')
    add(f"Space.{n}", f"Space({n})")
for n in (-1, 0, 1, 65, 255, 256, 65535, 65536):
    add(f"Chr.{n}", f"Chr({n})")
    add(f"ChrW.{n}", f"ChrW({n})")
V = ["0", "1", "-1", "0.5", "-0.5", "1.5", "2.5", "-1.5", "-2.5", "3.5", "100.5", "255.5", "32767.5", "2147483647.5", "1E+10", "1E+20", "123456.789", "-123456.789"]
for i, v in enumerate(V):
    for fn in ("Int", "Fix", "Abs", "Sgn", "CInt", "CLng", "CByte", "CDbl", "CSng", "CCur", "CStr", "CBool", "Hex", "Oct", "Sqr", "Exp", "Log", "Atn", "Sin", "Cos", "Tan", "Rnd2"):
        if fn == "Rnd2": continue
        add(f"{fn}.{i}", f"{fn}(vD{i})")
    for n in (-1, 0, 1, 2):
        add(f"Round.{i}.{n}", f"Round(vD{i}, {n})")
    add(f"Format.{i}", f'Format(vD{i}, "0.00")')
    add(f"FormatG.{i}", f'Format(vD{i}, "General Number")')
    add(f"FormatP.{i}", f'Format(vD{i}, "#,##0.0")')
    add(f"FormatC.{i}", f'Format(vD{i}, "0.0E+00")')
    add(f"FormatPct.{i}", f'Format(vD{i}, "0%")')
    add(f"Val.{i}", f"Val(CStr(vD{i}))")
for i, s in enumerate(S):
    for fn in ("CInt", "CLng", "CDbl", "CBool", "CByte", "CCur", "CSng", "CStr", "Val", "IsNumeric", "IsDate", "IsEmpty", "IsNull"):
        add(f"{fn}S.{i}", f"{fn}(vA{i})")
for a in ("1", "2", "0", "-1", "255", "256"):
    for b in ("1", "2", "0", "-1", "3"):
        add(f"And.{a}.{b}", f"{a} And {b}")
        add(f"Or.{a}.{b}", f"{a} Or {b}")
        add(f"Xor.{a}.{b}", f"{a} Xor {b}")
        add(f"Imp.{a}.{b}", f"{a} Imp {b}")
        add(f"Eqv.{a}.{b}", f"{a} Eqv {b}")
for x in ("True", "False", "Null"):
    for y in ("True", "False", "Null"):
        add(f"AndB.{x}.{y}", f"{x} And {y}")
        add(f"OrB.{x}.{y}", f"{x} Or {y}")
for d in ("#1/1/2000#", "#12/31/1999 11:59:59 PM#", "#2/29/2000#", "#3/1/1900#"):
    k = d.strip("#").replace("/", "_").replace(" ", "").replace(":", "")
    for fn in ("Year", "Month", "Day", "Hour", "Minute", "Second", "Weekday"):
        add(f"{fn}.{k}", f"{fn}({d})")
    add(f"CDbl.{k}", f"CDbl({d})")
    for iv in ("d", "m", "yyyy", "h", "n", "s", "w", "ww", "q"):
        add(f"DateAdd.{k}.{iv}", f"CDbl(DateAdd(\"{iv}\", 13, {d}))")
        add(f"DateDiff.{k}.{iv}", f"DateDiff(\"{iv}\", {d}, #6/15/2005 8:30:00 AM#)")
        add(f"DatePart.{k}.{iv}", f"DatePart(\"{iv}\", {d})")
lines = ["Function Rz(e As Long, val As Variant) As String\n  If e <> 0 Then\n    Rz = \"ERR\" & e\n  ElseIf IsNull(val) Then\n    Rz = \"Null\"\n  ElseIf VarType(val) = vbDate Then\n    Rz = \"Date:\" & Trim$(Str$(CDbl(val)))\n  Else\n    Rz = TypeName(val) & \":\" & CStr(val)\n  End If\nEnd Function\n"]
decl = "Dim r As Variant, out As String, " + ", ".join(f"vA{i} As Variant" for i in range(len(S))) + ", " + ", ".join(f"vD{i} As Variant" for i in range(len(V))) + "\n"
init = "  vA0 = \"\": vA1 = \"a\": vA2 = \"abc\": vA3 = \"Hello World\": vA4 = \"  pad  \": vA5 = \"12\": vA6 = \"-3.5\": vA7 = \"1e3\": vA8 = \"&H1F\": vA9 = \"abc abc\": vA10 = Null: vA11 = Empty\n  " + ": ".join(f"vD{i} = {v}" for i, v in enumerate(V)) + "\n"
names = []
for n in range(0, len(exprs), 100):
    nm = f"L{n // 100:03d}"
    names.append(nm)
    lines.append(f"Function {nm}() As String\n  {decl}  On Error Resume Next\n{init}")
    for id_, e in exprs[n:n + 100]:
        lines.append(f'  Err.Clear: r = Empty: r = {e}: out = out & "{id_}=" & Rz(Err.Number, r) & vbLf\n')
    lines.append(f"  {nm} = out\nEnd Function\n")
lines.append("Public Function RunAll() As String\n  Dim s As String\n" + "".join(f"  s = s & {nm}()\n" for nm in names) + "  RunAll = s\nEnd Function\n")
sys.stdout.write("".join(lines).replace("\n", "\r\n"))
sys.stderr.write(f"{len(exprs)} cases\n")
