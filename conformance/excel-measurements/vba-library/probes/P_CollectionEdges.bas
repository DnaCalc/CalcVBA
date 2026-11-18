Attribute VB_Name = "P_CollectionEdges"
Option Explicit
' UNMEASURED selector coercion and mutation order for a future live enumerator.
Private Const N As Long = 12
Private Function E(ByVal mode As Long) As String
    Dim c As New Collection, x As Variant, n As Long, s As String
    c.Add 1: c.Add 2: c.Add 3
    For Each x In c
        n = n + 1: s = s & Sh(x) & ";"
        If n = 1 Then
            Select Case mode
                Case 1: c.Add 99, , 1
                Case 2: c.Add 99, , 2
                Case 3: c.Remove 1: c.Add 99
                Case 4: c.Remove 2
            End Select
        End If
        If n > 10 Then Exit For
    Next
    E = s & "count=" & c.Count
End Function
Private Function T(ByVal k As Long) As String
    Dim v As Variant, c As New Collection
    On Error GoTo fail
    Select Case k
        Case 1: v = c.Item(0)
        Case 2: v = c.Item("missing")
        Case 3: c.Remove Null: v = c.Count
        Case 4: c.Add 1, , Null: v = c.Count
        Case 5: c.Add 1, , , 0: v = c.Count
        Case 6: c.Add 1: v = c.Item(CVErr(5))
        Case 7: c.Add 1: v = c.Item(Array(1))
        Case 8: c.Add 1: v = c.Item(Nothing)
        Case 9: v = E(1)
        Case 10: v = E(2)
        Case 11: v = E(3)
        Case 12: v = E(4)
    End Select
    T = "CollectionEdges." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "CollectionEdges." & k & "=ERR" & Err.Number
End Function
Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
