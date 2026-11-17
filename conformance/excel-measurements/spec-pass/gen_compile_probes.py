"""Writes the one-case probe modules for spec-pass (dc-ez3rd): lexical forms, date literals,
conditional compilation, Option rules, enums and declarations that may not compile in one or
both engines, so each is its own module (run-inproc-probe.ps1 -PerModule).

python -I gen_compile_probes.py   (writes Cmp*.bas beside this file)
"""

import os

HERE = os.path.dirname(os.path.abspath(__file__))

SHOW = (
    "Private Function Sh(ByVal v As Variant) As String\n"
    "    If VarType(v) = vbDate Then Sh = \"Date:\" & CStr(CDbl(v)) Else Sh = TypeName(v) & \":\" & CStr(v)\n"
    "End Function\n"
)


def run_value(key, expr, decl=""):
    """A module whose Run returns key=Sh(expr)."""
    return (
        f"{decl}{SHOW}"
        "Public Function Run() As String\n"
        "    On Error GoTo fail\n"
        f"    Run = \"{key}=\" & Sh({expr}) & vbLf: Exit Function\n"
        "fail:\n"
        f"    Run = \"{key}=ERR\" & Err.Number & vbLf\n"
        "End Function\n"
    )


MODULES = {
    # §3.2 a comment ending in " _" continues onto the next line
    "CmpCommentCont": (
        "Public Function Run() As String\n"
        "    Dim s As String\n"
        "    s = \"a\" ' note _\n"
        "    s = \"b\"\n"
        "    Run = \"commentCont=\" & s & vbLf\n"
        "End Function\n"
    ),
    # §3.3.5.3 the LongLong type character
    "CmpCaret": run_value("caret", "x", "Dim x^\n") .replace(
        "    On Error GoTo fail\n", "    On Error GoTo fail\n    x = 5\n"),
    # §3.3.2 number token forms
    "CmpNumD": run_value("numD", "1D3"),
    "CmpNumDot": run_value("numDot", ".5"),
    "CmpNumOct": run_value("numOct", "&17"),
    "CmpNumBig": run_value("numBig", "99999999999999999999"),
    # §3.3.3 date literals
    "CmpDate2Part": run_value("date2Part", "(Year(#1/15#) = Year(Date)) & \" \" & Month(#1/15#) & \"/\" & Day(#1/15#)"),
    "CmpDateMonYear": run_value("dateMonYear", "#3/2020#"),
    "CmpDateYY": run_value("dateYY", "#1/1/99#"),
    "CmpDateYY29": run_value("dateYY29", "#1/1/29#"),
    "CmpDateDash": run_value("dateDash", "#2-3-2026#"),
    "CmpDateYMD2": run_value("dateYMD2", "#99/12/31#"),
    "CmpDate3PM": run_value("date3PM", "#3 PM#"),
    "CmpDateAttachedAM": run_value("dateAttachedAM", "#8:30AM#"),
    "CmpDate13PM": run_value("date13PM", "#1/1/2026 13:00 PM#"),
    "CmpDateNoonPM": run_value("dateNoonPM", "#1/1/2026 0:00 PM#"),
    # §3.4 conditional compilation
    "CmpCcCont": (
        "#Const A = 1 _\n    + 1\n"
        "#If A = 2 Then\nPrivate Const R = \"two\"\n#Else\nPrivate Const R = \"other\"\n#End If\n"
        + run_value("ccCont", "R")
    ),
    "CmpCcIfCont": (
        "#If True _\n    Then\nPrivate Const R = \"yes\"\n#Else\nPrivate Const R = \"no\"\n#End If\n"
        + run_value("ccIfCont", "R")
    ),
    "CmpCcEndIf": (
        "#If True Then\nPrivate Const R = \"yes\"\n#EndIf\n" + run_value("ccEndIf", "R")
    ),
    "CmpCcSuffix": (
        "#Const X% = 1\n#If X Then\nPrivate Const R = \"yes\"\n#Else\nPrivate Const R = \"no\"\n#End If\n"
        + run_value("ccSuffix", "R")
    ),
    "CmpCcString": (
        "#Const S = \"x\"\n#If S Then\nPrivate Const R = \"yes\"\n#Else\nPrivate Const R = \"no\"\n#End If\n"
        + run_value("ccString", "R")
    ),
    "CmpCcVba6": (
        "#If VBA6 Then\nPrivate Const R = \"vba6\"\n#Else\nPrivate Const R = \"not\"\n#End If\n"
        + run_value("ccVba6", "R")
    ),
    "CmpCcDup": (
        "#If Win64 Then\n#Const P = 8\n#Else\n#Const P = 4\n#End If\n"
        "#If P = 8 Then\nPrivate Const R = \"8\"\n#Else\nPrivate Const R = \"4\"\n#End If\n"
        + run_value("ccDup", "R")
    ),
    # §5.2.1 option directives
    "CmpOptBase2": "Option Base 2\n" + run_value("optBase2", "1"),
    "CmpOptBase1Amp": "Option Base 1&\nDim a(3)\n" + run_value("optBase1Amp", "LBound(a)"),
    "CmpOptExplicit2": "Option Explicit\nOption Explicit\n" + run_value("optExplicit2", "1"),
    "CmpOptCompareFoo": "Option Compare Foo\n" + run_value("optCompareFoo", "(\"a\" = \"A\")"),
    # §5.2.3.4 enums
    "CmpEnumDbl": "Enum E1\n    A = 1.5\n    B\nEnd Enum\n" + run_value("enumDbl", "A & \" \" & B"),
    "CmpEnumOver": "Enum E2\n    A = 2147483647\n    B\nEnd Enum\n" + run_value("enumOver", "A"),
    "CmpEnumFwd": "Enum E3\n    X = Y\n    Y = 1\nEnd Enum\n" + run_value("enumFwd", "X"),
    # §5.3.1 property consistency
    "CmpPropMismatch": (
        "Private m As Long\n"
        "Property Get P(ByVal i As Long) As Long\n    P = m + i\nEnd Property\n"
        "Property Let P(ByRef i As Long, ByVal v As Long)\n    m = v\nEnd Property\n"
        + run_value("propMismatch", "P(1)")
    ),
    "CmpPropOptional": (
        "Private m As Long\n"
        "Property Get Q(Optional i As Long) As Long\n    Q = m\nEnd Property\n"
        "Property Let Q(i As Long, ByVal v As Long)\n    m = v\nEnd Property\n"
        + run_value("propOptional", "Q(1)")
    ),
    # §5.6.13.1 argument lists
    "CmpNamedThenPos": (
        "Private Function PA(a, ParamArray r()) As Long\n    PA = UBound(r)\nEnd Function\n"
        + run_value("namedThenPos", "PA(a:=1, 2)")
    ),
    "CmpCallByVal": (
        "Private Sub G(x)\n    x = 1\nEnd Sub\n"
        "Private Function H() As Variant\n    Dim k\n    G ByVal k\n    H = IsEmpty(k)\nEnd Function\n"
        + run_value("callByVal", "H()")
    ),
    # §5.6.9.3.6 & over a declared Byte() (moved out of SpecOps: VBE refuses it at compile)
    "CmpCatBytes": "Dim bb() As Byte\n" + run_value("catBytes", "bb & \"!\"").replace(
        "    On Error GoTo fail\n", "    On Error GoTo fail\n    bb = \"AB\"\n"),
    # §5.3.1.11 a Variant array element passed to a ByRef Long (moved out of SpecExpr)
    "CmpByRefVariantElem": (
        "Private Sub IncL(x As Long)\n    x = x + 1\nEnd Sub\n"
        "Private Function H() As Variant\n    Dim vv As Variant\n    vv = Array(1, 2)\n"
        "    IncL vv(0)\n    H = vv(0)\nEnd Function\n"
        + run_value("byRefVariantElem", "H()")
    ),
    # §5.4.2.3 GoTo into a For body
    "CmpGotoFor": (
        "Private Function H() As Long\n    Dim i As Long, n As Long\n    GoTo L\n"
        "    For i = 1 To 3\nL:      n = n + 1\n    Next\n    H = n\nEnd Function\n"
        + run_value("gotoFor", "H()")
    ),
}


def main():
    for name, body in MODULES.items():
        path = os.path.join(HERE, name + ".bas")
        with open(path, "w", encoding="utf-8", newline="\n") as out:
            out.write(f"Attribute VB_Name = \"{name}\"\n" + body)
    print(len(MODULES), "modules")


if __name__ == "__main__":
    main()
