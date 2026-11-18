Attribute VB_Name = "P_Switch"
Option Explicit
' Switch (Interaction; first True expression wins, all arguments evaluated, odd count is error 5).
Private cnt As Long
Private Const N As Long = 27

Private Function Tick(ByVal x As Variant) As Variant
    cnt = cnt + 1
    Tick = x
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    Dim z As Long
    On Error GoTo fail
    Select Case k
        Case 1: v = Switch(True, "a", False, "b")     ' first true
        Case 2: v = Switch(False, "a", True, "b")     ' second true
        Case 3: v = Switch(True, "a", True, "b")      ' both true: first wins
        Case 4: v = Switch(False, "a", False, "b")    ' none true: Null
        Case 5: v = Switch(False, 1)                  ' single false pair: Null
        Case 6: v = Switch(True, 1)                   ' single true pair
        Case 7: v = Switch(True)                      ' odd count: error 5
        Case 8: v = Switch(True, 1, False)            ' odd count: error 5
        Case 9: v = Switch(False, 1, False, 2, True)  ' odd count: error 5
        Case 10: v = Switch(Null, 1, True, 2)          ' Null expression
        Case 11: v = Switch(Empty, 1, True, 2)         ' Empty expression
        Case 12: v = Switch(0, "a", 1, "b")            ' numeric expressions
        Case 13: v = Switch(2, "a")                    ' non-zero is True
        Case 14: v = Switch("True", "a")               ' string expression converts
        Case 15: v = Switch("abc", 1)                  ' error 13
        Case 16: v = Switch(True, Null)                ' Null value
        Case 17: v = Switch(True, 1, False, "a")       ' result type
        Case 18: v = Switch(False, 1, True, "a")
        Case 19: cnt = 0: v = Switch(True, Tick("a"), True, Tick("b")) & "|" & cnt     ' all arguments evaluated
        Case 20: cnt = 0: v = Switch(False, Tick("a"), True, Tick("b"), True, Tick("c")) & "|" & cnt
        Case 21: v = Switch(True, 1, True, 1 \ z)      ' unselected failing value still raises: error 11
        Case 22: v = Switch(1 \ z = 0, 1, True, 2)     ' failing expression: error 11
        Case 23: v = Switch(1 = 1, "x", 1 = 2, "y")
        Case 24: v = Switch(False, 1, False, 2, False, 3, True, 4)   ' four pairs
        Case 25: v = Switch(Array(1), 1)               ' array expression: error 13
        Case 26: v = Switch(True, CVErr(5))            ' error value
        Case 27: v = TypeName(Switch(True, 1))
    End Select
    T = "Switch." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Switch." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
