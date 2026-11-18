Attribute VB_Name = "P_Command"
Option Explicit
' Command / Command$ (Interaction; the command line of the host, only types and relations printed).
Private Const N As Long = 8

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = TypeName(Command)                             ' String
        Case 2: v = TypeName(Command$)                            ' String
        Case 3: v = Command = Command$                            ' both forms agree
        Case 4: v = Len(Command) >= 0                             ' length
        Case 5: v = VarType(Command)
        Case 6: v = TypeName(Command())                           ' empty parentheses
        Case 7: v = Command() = Command
        Case 8: v = Len(Command & "") = Len(Command)              ' concatenation with empty
    End Select
    T = "Command." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Command." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
