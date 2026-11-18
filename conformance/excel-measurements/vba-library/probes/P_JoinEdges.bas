Attribute VB_Name = "P_JoinEdges"
Option Explicit
' UNMEASURED empty array kinds, error precedence and raw UTF16 preservation.
Private Const N As Long = 8
Private Function T(ByVal k As Long) As String
    Dim v As Variant, va() As Variant, la() As Long, sa() As String, m(0 To 0, 0 To 0) As String
    On Error GoTo fail
    Select Case k
        Case 1: v = Join(va, ",")
        Case 2: v = Join(la, ",")
        Case 3: v = Join(sa, Null)
        Case 4: v = Join(m, Null)
        Case 5: v = Join(Null, Null)
        Case 6: v = Join(Array("a", "b"), CVErr(5))
        Case 7: v = Join(Array(ChrW(&HD800), "b"), ",")
        Case 8: v = Join(Array("a", "b"), ChrW(&HDC00))
    End Select
    T = "JoinEdges." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "JoinEdges." & k & "=ERR" & Err.Number
End Function
Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
