Attribute VB_Name = "P_RegExp"
Option Explicit
' VBA.RegExp class: Pattern, Global, IgnoreCase, Multiline, Test, Execute, Replace.
Private Const N As Long = 44

Private Function Mk(ByVal p As String, Optional ByVal g As Boolean = False, Optional ByVal ic As Boolean = False, Optional ByVal ml As Boolean = False) As VBA.RegExp
    Dim r As VBA.RegExp
    Set r = New VBA.RegExp
    r.Pattern = p
    r.Global = g
    r.IgnoreCase = ic
    r.Multiline = ml
    Set Mk = r
End Function

Private Function Ms(ByVal mc As Object) As String
    Dim i As Long, s As String, m As Object
    s = "n=" & mc.Count & ":"
    For i = 0 To mc.Count - 1
        Set m = mc.Item(i)
        s = s & "[" & m.FirstIndex & "," & m.Length & "," & Esc(m.Value) & "]"
    Next
    Ms = s
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant, r As VBA.RegExp, mc As Object, sm As Object
    On Error GoTo fail
    Select Case k
        Case 1: Set r = New VBA.RegExp: v = Sh(r.Pattern) & Sh(r.Global) & Sh(r.IgnoreCase) & Sh(r.Multiline)    ' defaults
        Case 2: v = Mk("b+").Test("aabbbcc")                                   ' Test true
        Case 3: v = Mk("x").Test("aabbbcc")                                    ' Test false
        Case 4: v = Mk("abc").Test("xABCx")                                    ' case sensitive by default
        Case 5: v = Mk("abc", , True).Test("xABCx")                            ' IgnoreCase
        Case 6: v = Ms(Mk("a\d").Execute("a1a2a3"))                            ' Global false: one match
        Case 7: v = Ms(Mk("a\d", True).Execute("a1a2a3"))                      ' Global true
        Case 8: v = Mk("^b").Test("a" & vbLf & "b")                            ' ^ without Multiline
        Case 9: v = Mk("^b", , , True).Test("a" & vbLf & "b")                  ' ^ with Multiline
        Case 10: v = Mk("a$", , , True).Test("a" & vbLf & "b")                 ' $ with Multiline
        Case 11: v = Mk("a$", , , True).Test("a" & vbCrLf & "b")               ' $ before CRLF
        Case 12: v = Mk("(\w+) (\w+)").Replace("hello world", "$2 $1")         ' $1 $2
        Case 13: v = Mk("o", True).Replace("foo", "[$&]")                      ' $&
        Case 14: v = Mk("b").Replace("abc", "<$`|$'>")                         ' $` and $'
        Case 15: v = Mk("b").Replace("abc", "$$")                              ' $$
        Case 16: v = Mk("(b)").Replace("abc", "$3")                            ' group that does not exist
        Case 17: v = Mk("\d").Replace("a1a2", "X")                             ' Global false replaces first only
        Case 18: v = Mk("\d", True).Replace("a1a2", "X")                       ' Global true replaces all
        Case 19: v = Mk("z").Replace("abc", "X")                               ' no match: source returned
        Case 20: v = Mk("b").Replace("abc", 5)                                 ' numeric replaceVar
        Case 21: v = Mk("b").Replace("abc", Null)                              ' Null replaceVar
        Case 22: v = Mk("b", True).Replace("abcb", "")                         ' empty replacement
        Case 23: Set mc = Mk("b+").Execute("aabbbcc"): Set sm = mc.Item(0): v = sm.Value & "," & sm.FirstIndex & "," & sm.Length    ' Match properties
        Case 24: Set mc = Mk("(\d+)-(\d+)").Execute("10-20"): v = mc.Item(0).SubMatches.Count & "," & mc.Item(0).SubMatches(0) & "," & mc.Item(0).SubMatches.Item(1)    ' SubMatches
        Case 25: Set mc = Mk("(a)|(b)").Execute("b"): v = Sh(mc.Item(0).SubMatches(0)) & Sh(mc.Item(0).SubMatches(1))    ' group that did not take part
        Case 26: v = Mk("(").Test("a")                                         ' invalid pattern: unclosed group
        Case 27: v = Mk("[").Test("a")                                         ' invalid pattern: unclosed class
        Case 28: v = Mk("*").Test("a")                                         ' invalid pattern: nothing to repeat
        Case 29: v = Mk("a{2,1}").Test("aa")                                   ' invalid quantifier range
        Case 30: v = Ms(Mk("a(?=b)", True).Execute("ab ac"))                   ' lookahead
        Case 31: v = Ms(Mk("a(?!b)", True).Execute("ab ac"))                   ' negative lookahead
        Case 32: v = Ms(Mk("(?<=a)b", True).Execute("ab cb"))                  ' lookbehind (may be unsupported)
        Case 33: v = Ms(Mk("[A-C]+\d\s\w\b", True, True).Execute("abc1 d x"))  ' classes and \b with IgnoreCase
        Case 34: v = Ms(Mk("(a)\1", True).Execute("aa ab aaaa"))               ' backreference
        Case 35: v = Ms(Mk("<.+?>", True).Execute("<a><b>"))                   ' non-greedy
        Case 36: v = Mk("a.b").Test("a" & vbLf & "b")                          ' dot does not match newline
        Case 37: v = Mk(ChrW(233), , True).Test(ChrW(201))                     ' IgnoreCase on e-acute
        Case 38: v = Mk("^\w$").Test(ChrW(233))                                ' \w with e-acute
        Case 39: v = Ms(Mk(".", True).Execute(ChrW(&HD83D) & ChrW(&HDE00)))    ' surrogate pair: dot matches one unit or two
        Case 40: v = Mk("é").Test(ChrW(233))                              ' \u escape
        Case 41: v = Ms(Mk("", True).Execute("abc"))                           ' empty pattern, Global
        Case 42: v = Mk("^$").Test("")                                         ' empty source
        Case 43: v = Mk("a").Replace(sourceString:="cat", replaceVar:="o") & "," & Mk("a").Test(sourceString:="cat") & "," & Mk("a").Execute(sourceString:="cat").Count    ' named arguments
        Case 44: Set r = Mk("a", True): v = r.Test("a") & "," & r.Test("a") & "," & TypeName(r) & "," & TypeName(r.Execute("a"))    ' repeated Test with Global; type names
    End Select
    T = "RegExp." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "RegExp." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
