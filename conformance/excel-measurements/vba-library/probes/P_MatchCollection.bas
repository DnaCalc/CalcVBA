Attribute VB_Name = "P_MatchCollection"
Option Explicit
' VBA.MatchCollection class: Count, Item, For Each, default member.
Private Const N As Long = 20

Private Function Mk(ByVal p As String, Optional ByVal g As Boolean = False) As VBA.RegExp
    Dim r As VBA.RegExp
    Set r = New VBA.RegExp
    r.Pattern = p
    r.Global = g
    Set Mk = r
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant, mc As Object, mc2 As Object, x As Variant, s As String, r As VBA.RegExp
    On Error GoTo fail
    Select Case k
        Case 1: Set mc = Mk("a", True).Execute("banana"): v = mc.Count                       ' Count
        Case 2: Set mc = Mk("a", True).Execute("banana"): v = TypeName(mc) & "," & TypeName(mc.Count)
        Case 3: Set mc = Mk("a", True).Execute("banana"): v = mc.Item(0).FirstIndex          ' first, zero-based
        Case 4: Set mc = Mk("a", True).Execute("banana"): v = mc.Item(2).FirstIndex          ' last
        Case 5: Set mc = Mk("a", True).Execute("banana"): v = mc.Item(3).FirstIndex          ' index = Count: error
        Case 6: Set mc = Mk("a", True).Execute("banana"): v = mc.Item(-1).FirstIndex         ' negative index
        Case 7: Set mc = Mk("a", True).Execute("banana"): v = mc.Item(1.5).FirstIndex        ' Double index rounds
        Case 8: Set mc = Mk("a", True).Execute("banana"): v = mc.Item("1").FirstIndex        ' string digits
        Case 9: Set mc = Mk("a", True).Execute("banana"): v = mc.Item(Null).FirstIndex       ' Null index
        Case 10: Set mc = Mk("a", True).Execute("banana"): For Each x In mc: s = s & x.FirstIndex & ";": Next: v = s    ' For Each order
        Case 11: Set mc = Mk("a", True).Execute("banana"): For Each x In mc: s = s & TypeName(x) & ";": Next: v = s    ' loop-variable type
        Case 12: Set mc = Mk("z", True).Execute("banana"): For Each x In mc: s = s & "x": Next: v = mc.Count & "/" & s    ' no matches
        Case 13: Set mc = Mk("a", True).Execute("banana"): v = mc(1).FirstIndex             ' default member
        Case 14: Set mc = Mk("a").Execute("banana"): v = mc.Count                            ' Global false: one match
        Case 15: Set mc = Mk("aa", True).Execute("aaaaa"): v = mc.Count                      ' non-overlapping
        Case 16: Set r = Mk("a", True): Set mc = r.Execute("banana"): Set mc2 = r.Execute("a"): v = mc.Count & "," & mc2.Count    ' earlier collection unchanged
        Case 17: Set r = Mk("a", True): Set mc = r.Execute("banana"): r.Pattern = "n": v = mc.Count & "," & r.Execute("banana").Count    ' pattern change after Execute
        Case 18: Set mc = Mk("", True).Execute("ab"): v = mc.Count                           ' empty pattern
        Case 19: Set mc = Mk("a", True).Execute("banana"): v = mc.Item(Index:=1).Value       ' named argument
        Case 20: Set mc = Mk("a", True).Execute("banana"): Set mc2 = mc: v = (mc Is mc2) & "," & TypeName(mc.Item(0))    ' identity; item class
    End Select
    T = "MatchCollection." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "MatchCollection." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
