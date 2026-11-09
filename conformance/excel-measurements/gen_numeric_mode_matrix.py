#!/usr/bin/env python3
"""Generates numeric-mode-matrix.vba: Excel VBA probe for the NumericMode decision matrix.

Every case runs under On Error Resume Next and records `id=<TypeName>:<value>` or `id=ERR<n>`.
The same module runs on Excel (measurement) and on the engines (test).
"""
import sys

TYPES = [
    ("Byte", "b", ["0", "200", "255"]),
    ("Integer", "i", ["-32768", "1", "32767"]),
    ("Long", "l", ["-2147483648", "65536", "2147483647"]),
    ("LongLong", "q", ["-9223372036854775807^", "3", "9223372036854775807^"]),
    ("Single", "s", ["0.5", "16777216", "3.0E+38"]),
    ("Double", "d", ["0.5", "2.5", "1.0E+308"]),
    ("Currency", "c", ["-0.5@", "1.5@", "922337203685477.5807@"]),
    ("Variant", "v", ["Empty", "7", "2147483647"]),
]
VARIANT_FORMS = [  # (suffix, assignment text)
    ("vInt", "vA = 32767"),
    ("vLng", "vA = 2147483647"),
    ("vDbl", "vA = 2.5"),
    ("vCur", "vA = 1.5@"),
    ("vStr", 'vA = "12"'),
    ("vEmp", "vA = Empty"),
]
OPS = ["+", "-", "*", "/", "\\", "Mod", "^"]
CONV = ["CByte", "CInt", "CLng", "CLngLng", "CSng", "CDbl", "CCur"]
CONV_INPUTS = ["-0.5", "0.5", "1.5", "2.5", "255.5", "256", "32767.5", "32768", "2147483647.5",
               "2147483648", "4.5E+18", "9.3E+18", "-1", "1.0E+38", "3.5E+38", "1.0E+300",
               "922337203685477.58", "922337203685478"]
CONV_STR = ['"12"', '"1.5"', '"2.5"', '"-0.5"', '" 7 "', '"1e3"', '"&H10"', '"abc"', '""', '"32768"', '"99999999999"']

def chunks(lst, n):
    for k in range(0, len(lst), n):
        yield lst[k:k + n]

cases = []  # (id, vba line)
def case(id_, setup, expr):
    cases.append((id_, f'Err.Clear: {setup}: r = {expr}: out = out & "{id_}=" & Rz(Err.Number, r) & vbLf'))

def assign(prefix, setup, expr, target):
    cases.append((prefix, f'Err.Clear: {setup}: {target} = {expr}: out = out & "{prefix}=" & Rz(Err.Number, {target}) & vbLf'))

TV = {n: (p, vals) for n, p, vals in TYPES}
# binary operations over typed operand pairs
for lt, lp, lvals in TYPES:
    for rt, rp, rvals in TYPES:
        for li, lv in enumerate(lvals):
            for ri, rv in enumerate(rvals):
                if lt == "Variant" and lv == "Empty" and rt == "Variant" and rv == "Empty":
                    continue
                for op in OPS:
                    opname = {"+": "add", "-": "sub", "*": "mul", "/": "div", "\\": "idiv", "Mod": "mod", "^": "pow"}[op]
                    id_ = f"bin.{lt}{li}.{opname}.{rt}{ri}"
                    setup = f"{lp}A = {lv}: {rp}B = {rv}" if lp != rp else f"{lp}A = {lv}: {rp}B = {rv}"
                    case(id_, setup, f"{lp}A {op} {rp}B")
# Variant forms against typed operands
for form, asg in VARIANT_FORMS:
    for rt, rp, rvals in TYPES[:7]:
        for ri, rv in enumerate(rvals):
            for op in OPS:
                opname = {"+": "add", "-": "sub", "*": "mul", "/": "div", "\\": "idiv", "Mod": "mod", "^": "pow"}[op]
                case(f"var.{form}.{opname}.{rt}{ri}", f"{asg}: {rp}B = {rv}", f"vA {op} {rp}B")
                case(f"var.{rt}{ri}.{opname}.{form}", f"{asg}: {rp}B = {rv}", f"{rp}B {op} vA")
# unary minus and Not
for lt, lp, lvals in TYPES:
    for li, lv in enumerate(lvals):
        case(f"neg.{lt}{li}", f"{lp}A = {lv}", f"-{lp}A")
        case(f"not.{lt}{li}", f"{lp}A = {lv}", f"Not {lp}A")
        case(f"abs.{lt}{li}", f"{lp}A = {lv}", f"Abs({lp}A)")
# typed assignment of an expression (overflow on store versus on compute)
ASSIGN_TARGETS = [("Byte", "bR"), ("Integer", "iR"), ("Long", "lR"), ("LongLong", "qR"), ("Single", "sR"), ("Double", "dR"), ("Currency", "cR")]
for tt, tv in ASSIGN_TARGETS:
    for lt, lp, lvals in TYPES[:7]:
        # Excel rejects LongLong -> Byte/Integer/Long/Currency stores at compile time (type mismatch);
        # those shapes are covered by the compile-only cases in numeric-mode-compile.tsv.
        if lt == 'LongLong' and tt in ('Byte', 'Integer', 'Long', 'Currency'):
            continue
        for li, lv in enumerate(lvals):
            assign(f"store.{tt}.from.{lt}{li}", f"{lp}A = {lv}", f"{lp}A", tv)
            assign(f"store.{tt}.add1.{lt}{li}", f"{lp}A = {lv}", f"{lp}A + 1", tv)
            assign(f"store.{tt}.mul2.{lt}{li}", f"{lp}A = {lv}", f"{lp}A * 2", tv)
# conversions
for cv in CONV:
    for k, x in enumerate(CONV_INPUTS):
        case(f"conv.{cv}.d{k}", f"dA = {x}", f"{cv}(dA)")
    for k, x in enumerate(CONV_STR):
        case(f"conv.{cv}.s{k}", f"vA = {x}", f"{cv}(vA)")
    for form, asg in VARIANT_FORMS:
        case(f"conv.{cv}.{form}", asg, f"{cv}(vA)")
# comparisons
for lt, lp, lvals in TYPES[:7]:
    for rt, rp, rvals in TYPES[:7]:
        for li, lv in enumerate(lvals):
            for ri, rv in enumerate(rvals):
                for op, nm in (("=", "eq"), ("<", "lt")):
                    case(f"cmp.{lt}{li}.{nm}.{rt}{ri}", f"{lp}A = {lv}: {rp}B = {rv}", f"({lp}A {op} {rp}B)")
# integer functions
for fn in ("Int", "Fix", "Sgn", "CInt", "Round"):
    for lt, lp, lvals in TYPES[4:7]:
        for li, lv in enumerate(lvals):
            case(f"fn.{fn}.{lt}{li}", f"{lp}A = {lv}", f"{fn}({lp}A)")

DECL = "Dim bA As Byte, bB As Byte, iA As Integer, iB As Integer, lA As Long, lB As Long, qA As LongLong, qB As LongLong, sA As Single, sB As Single, dA As Double, dB As Double, cA As Currency, cB As Currency, vA As Variant, vB As Variant, r As Variant, out As String\n  Dim bR As Byte, iR As Integer, lR As Long, qR As LongLong, sR As Single, dR As Double, cR As Currency\n"
lines = []
lines.append('Function Rz(e As Long, val As Variant) As String\n  If e <> 0 Then\n    Rz = "ERR" & e\n  Else\n    Rz = TypeName(val) & ":" & Trim$(Str$(val))\n  End If\nEnd Function\n')
names = []
only = sys.argv[1] if len(sys.argv) > 1 else None
if only:
    cases = [c for c in cases if c[0].startswith(only)]
for n, ch in enumerate(chunks(cases, 120)):
    nm = f"P{n:03d}"
    names.append(nm)
    lines.append(f"Function {nm}() As String\n  {DECL}  On Error Resume Next\n")
    for _, ln in ch:
        lines.append("  " + ln + "\n")
    lines.append(f"  {nm} = out\nEnd Function\n")
lines.append("Public Function RunAll() As String\n  Dim s As String\n")
for nm in names:
    lines.append(f"  s = s & {nm}()\n")
lines.append("  RunAll = s\nEnd Function\n")
sys.stdout.write("".join(lines).replace("\n", "\r\n"))
sys.stderr.write(f"{len(cases)} cases, {len(names)} procedures\n")
