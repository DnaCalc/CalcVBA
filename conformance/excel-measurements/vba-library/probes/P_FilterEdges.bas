Attribute VB_Name = "P_FilterEdges"
Option Explicit
' UNMEASURED argument precedence, coercions and empty source behavior.
Private Const N As Long = 18
Private Function T(ByVal k As Long) As String
    Dim v As Variant, sa() As String, la() As Long, m(0 To 0, 0 To 0) As String
    On Error GoTo fail
    Select Case k
        Case 1: v = Filter(Array("a", "b"), "a", Null)
        Case 2: v = Filter(Array("a", "b"), "a", "True")
        Case 3: v = Filter(Array("a", "b"), "a", "False")
        Case 4: v = Filter(Array("a", "b"), "a", "abc")
        Case 5: v = Filter(Array("a", "b"), "a", 0.5)
        Case 6: v = Filter(Array("a", "b"), "a", CVErr(5))
        Case 7: v = Filter(sa, Null)
        Case 8: v = Filter(sa, "a", Null)
        Case 9: v = Filter(Array(), Null)
        Case 10: v = Filter(la, "a")
        Case 11: v = Filter(m, Null)
        Case 12: v = Filter(Array("a", Empty), "")
        Case 13: v = Filter(Array(True, False), "True")
        Case 14: v = Filter(Array("a", CVErr(5)), "z")
        Case 15: v = Filter(Array("a", "b"), "a", True, Null)
        Case 16: v = Filter(Array("a", "A"), "a", True, 3)
        Case 17: v = Filter(Array(ChrW(&HD800), "a"), "")
        Case 18: v = Filter(Array("a", "b"), "a", False, -1)
    End Select
    T = "FilterEdges." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "FilterEdges." & k & "=ERR" & Err.Number
End Function
Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
