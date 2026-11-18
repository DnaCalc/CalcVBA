Attribute VB_Name = "P_RegExpEdges"
Option Explicit
' Additional dialect/argument probes. No Excel answers captured on the Linux lane yet.
Private Const N As Long = 32

Private Function T(ByVal k As Long) As String
    Dim r As RegExp, mc As MatchCollection, m As Match, sm As SubMatches
    Dim o As Object, v As Variant
    On Error GoTo fail
    Set r = New RegExp
    r.Pattern = "(a)(b)?"
    Select Case k
        Case 1: Set o = CreateObject("VBScript.RegExp"): v = TypeName(o) & "," & o.Test("")
        Case 2: Set mc = r.Execute("a"): Set m = mc(0): Set sm = m.SubMatches: v = TypeName(mc) & "," & TypeName(m) & "," & TypeName(sm)
        Case 3: v = r.Replace(replaceVar:="[$1]", sourceString:="a")
        Case 4: Set mc = r.Execute("a"): v = CallByName(mc, "Count", vbGet)
        Case 5: r.Pattern = ".": v = r.Test(ChrW(&H2028))
        Case 6: r.Pattern = ".": v = r.Test(ChrW(&H2029))
        Case 7: r.Pattern = "a$": v = r.Test("a" & vbLf)
        Case 8: r.Pattern = "a$": v = r.Test("a" & vbCr)
        Case 9: r.Pattern = "^b": r.Multiline = True: v = r.Test("a" & vbCr & "b")
        Case 10: r.Pattern = "^b": r.Multiline = True: v = r.Test("a" & ChrW(&H2028) & "b")
        Case 11: r.Pattern = "a": v = r.Replace("a", "$0/$00/$01/$10/$99")
        Case 12: v = r.Replace("a", "$0/$00/$01/$10/$99")
        Case 13: v = r.Replace("a", "$2")
        Case 14: r.Pattern = "(a)(b)(c)(d)(e)(f)(g)(h)(i)(j)(k)": v = r.Replace("abcdefghijk", "$11/$12/$01/$00")
        Case 15: r.Pattern = "a": r.Global = True: v = r.Replace("aa", "<$`|$'>")
        Case 16: v = r.Test(Null)
        Case 17: r.Pattern = Null: v = r.Pattern
        Case 18: r.Global = Null: v = r.Global
        Case 19: r.Global = 1.5: v = r.Global
        Case 20: r.Global = "True": v = r.Global
        Case 21: r.Pattern = "a": v = r.Replace("a", Empty)
        Case 22: r.Pattern = "a": v = r.Replace("a", CVErr(5))
        Case 23: r.Pattern = "(?<x>a)": v = r.Test("a")
        Case 24: r.Pattern = "(?i)a": v = r.Test("a")
        Case 25: r.Pattern = "a++": v = r.Test("aa")
        Case 26: r.Pattern = "a)": v = r.Test("a")
        Case 27: r.Pattern = "[z-a]": v = r.Test("a")
        Case 28: r.Pattern = "[]": v = r.Test("a")
        Case 29: r.Pattern = "\": v = r.Test("a")
        Case 30: r.Pattern = "(?=a)*": v = r.Test("a")
        Case 31: Set mc = r.Execute("a"): Set m = mc(0): v = (m.SubMatches Is m.SubMatches)
        Case 32: r.Pattern = "": r.Global = True: v = r.Replace("ab", "X")
    End Select
    T = "RegExpEdges." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "RegExpEdges." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
