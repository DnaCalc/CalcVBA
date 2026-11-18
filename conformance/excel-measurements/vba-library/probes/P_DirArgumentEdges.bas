Attribute VB_Name = "P_DirArgumentEdges"
Option Explicit
' UNMEASURED subtype edges for the Dir validation added by lane vl1.
' Print only the result type, never directory contents.
Private Const N As Long = 5
Private Function T(ByVal k As Long) As String
    Dim v As Variant, reset As String
    On Error GoTo fail
    reset = Dir("__calcvba_dir_argument_edges_no_such_entry__")
    Select Case k
        Case 1: v = TypeName(Dir(1))
        Case 2: v = TypeName(Dir(True))
        Case 3: v = TypeName(Dir(Empty))
        Case 4: v = TypeName(Dir("1"))
        Case 5: v = TypeName(Dir(Null))
    End Select
    T = "DirArgumentEdges." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "DirArgumentEdges." & k & "=ERR" & Err.Number
End Function
Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
