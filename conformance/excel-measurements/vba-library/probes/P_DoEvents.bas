Attribute VB_Name = "P_DoEvents"
Option Explicit
' DoEvents (Interaction; returns the number of open forms, an Integer).
Private Const N As Long = 10

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = DoEvents                                      ' value
        Case 2: v = TypeName(DoEvents)                            ' Integer
        Case 3: v = VarType(DoEvents)                             ' 2
        Case 4: DoEvents: v = "statement form ok"                 ' called as a statement
        Case 5: v = "compile-refused"
        Case 6: v = (DoEvents = 0)                                ' no forms open
        Case 7: v = CLng(DoEvents) + 1                            ' used in arithmetic
        Case 8: v = DoEvents + DoEvents                           ' two calls
        Case 9: v = DoEvents()                                    ' empty parentheses
        Case 10: v = IsNumeric(DoEvents)
    End Select
    T = "DoEvents." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "DoEvents." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
