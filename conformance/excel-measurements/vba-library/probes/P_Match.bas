Attribute VB_Name = "P_Match"
Option Explicit
' VBA.Match class (and its SubMatches collection): Value, FirstIndex, Length, SubMatches.
Private Const N As Long = 26

Private Function Mk(ByVal p As String, Optional ByVal g As Boolean = False, Optional ByVal ic As Boolean = False) As VBA.RegExp
    Dim r As VBA.RegExp
    Set r = New VBA.RegExp
    r.Pattern = p
    r.Global = g
    r.IgnoreCase = ic
    Set Mk = r
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant, m As Object, m2 As Object, mc As Object, sm As Object, x As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: Set m = Mk("b+").Execute("aabbbcc").Item(0): v = m.Value                    ' Value
        Case 2: Set m = Mk("b+").Execute("aabbbcc").Item(0): v = m.FirstIndex               ' FirstIndex is zero-based
        Case 3: Set m = Mk("b+").Execute("aabbbcc").Item(0): v = m.Length                   ' Length
        Case 4: Set m = Mk("b+").Execute("aabbbcc").Item(0): v = TypeName(m)                ' class name
        Case 5: Set m = Mk("b+").Execute("aabbbcc").Item(0): v = TypeName(m.Value) & "," & TypeName(m.FirstIndex) & "," & TypeName(m.Length) & "," & TypeName(m.SubMatches)
        Case 6: Set m = Mk("b+").Execute("aabbbcc").Item(0): v = m.SubMatches.Count          ' no groups
        Case 7: Set m = Mk("(\d+)-(\d+)").Execute("x10-20").Item(0): v = m.SubMatches.Count & "," & TypeName(m.SubMatches.Count)    ' two groups
        Case 8: Set m = Mk("(\d+)-(\d+)").Execute("x10-20").Item(0): v = m.SubMatches(0)     ' default Item
        Case 9: Set m = Mk("(\d+)-(\d+)").Execute("x10-20").Item(0): v = m.SubMatches.Item(1)    ' explicit Item
        Case 10: Set m = Mk("(\d+)-(\d+)").Execute("x10-20").Item(0): v = m.SubMatches(5)    ' out of range
        Case 11: Set m = Mk("(\d+)-(\d+)").Execute("x10-20").Item(0): v = m.SubMatches(-1)   ' negative index
        Case 12: Set m = Mk("(\d+)-(\d+)").Execute("x10-20").Item(0): For Each x In m.SubMatches: s = s & TypeName(x) & ":" & x & ";": Next: v = s    ' For Each over SubMatches
        Case 13: Set m = Mk("b+").Execute("aabbbcc").Item(0): v = m                          ' default member of Match (Value?)
        Case 14: Set m = Mk("x*").Execute("abc").Item(0): v = m.Length & "," & m.FirstIndex & "," & Esc(m.Value)    ' empty match
        Case 15: Set mc = Mk("ab", True).Execute("xx ab ab"): v = mc.Item(1).FirstIndex      ' FirstIndex of the second match
        Case 16: Set m = Mk("b").Execute(ChrW(&HD83D) & ChrW(&HDE00) & "b").Item(0): v = m.FirstIndex    ' UTF-16 units before the match
        Case 17: Set m = Mk("a.*").Execute("xa" & vbCr & "b").Item(0): v = m.Length & "," & Esc(m.Value)    ' dot and CR
        Case 18: Set m = Mk("abc", , True).Execute("xABCx").Item(0): v = m.Value             ' keeps the source case
        Case 19: Set m = Mk("((a)(b))").Execute("ab").Item(0): v = m.SubMatches(0) & "," & m.SubMatches(1) & "," & m.SubMatches(2)    ' nested group order
        Case 20: Set m = Mk("(a)|(b)").Execute("b").Item(0): v = m.SubMatches.Count & Sh(m.SubMatches(0))    ' group not taken
        Case 21: Set mc = Mk("a").Execute("aa"): Set m = mc.Item(0): Set m2 = mc.Item(0): v = (m Is m2)    ' same object each time?
        Case 22: Set m = Mk("(a)").Execute("a").Item(0): v = TypeName(m.SubMatches(0)) & "," & VarType(m.SubMatches(0))    ' SubMatch item type
        Case 23: Set m = Mk("(a)").Execute("a").Item(0): Set sm = m.SubMatches: v = TypeName(sm)    ' SubMatches class name
        Case 24: Set m = Mk("(a)?b").Execute("b").Item(0): v = Sh(m.SubMatches(0))           ' optional group absent
        Case 25: Set m = Mk(ChrW(233) & "|é").Execute(ChrW(233)).Item(0): v = AscW(m.Value) & "," & m.Length    ' Unicode value
        Case 26: Set m = Mk("(a)(b)(c)(d)(e)(f)(g)(h)(i)(j)(k)").Execute("abcdefghijk").Item(0): v = m.SubMatches.Count & m.SubMatches(10)    ' eleven groups
    End Select
    T = "Match." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Match." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
