Attribute VB_Name = "P_IIf"
Option Explicit
' IIf (Interaction; both branches are always evaluated).
Private cnt As Long
Private Const N As Long = 32

Private Function Tick(ByVal x As Variant) As Variant
    cnt = cnt + 1
    Tick = x
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    Dim va As Variant, c As Collection, o As Object, z As Long
    On Error GoTo fail
    Select Case k
        Case 1: v = IIf(True, 1, 2)                   ' true branch
        Case 2: v = IIf(False, 1, 2)                  ' false branch
        Case 3: v = IIf(1 = 1, "yes", "no")
        Case 4: v = IIf(1, "a", "b")                  ' non-zero number is True
        Case 5: v = IIf(0, "a", "b")
        Case 6: v = IIf(-1, "a", "b")
        Case 7: v = IIf(0.5, "a", "b")
        Case 8: v = IIf("True", "a", "b")             ' string condition converts
        Case 9: v = IIf("1", "a", "b")
        Case 10: v = IIf("0", "a", "b")
        Case 11: v = IIf("abc", "a", "b")              ' error 13
        Case 12: v = IIf("", "a", "b")                 ' error 13
        Case 13: v = IIf(Null, "a", "b")               ' Null condition: error 94
        Case 14: v = IIf(Empty, "a", "b")              ' Empty is False
        Case 15: v = IIf(Array(1), "a", "b")           ' array condition: error 13
        Case 16: v = IIf(True, 1, "a")                 ' result type Integer
        Case 17: v = IIf(False, 1, "a")                ' result type String
        Case 18: v = IIf(True, Null, 1)                ' Null branch
        Case 19: v = IIf(False, 1, Null)
        Case 20: v = IIf(True, Empty, 1)               ' Empty branch
        Case 21: v = IIf(True, 1, 1 \ z)               ' false branch evaluated: error 11
        Case 22: v = IIf(False, 1 \ z, 2)              ' true branch evaluated: error 11
        Case 23: cnt = 0: v = IIf(True, Tick("a"), Tick("b")) & "|" & cnt     ' both evaluated
        Case 24: cnt = 0: v = IIf(False, Tick("a"), Tick("b")) & "|" & cnt
        Case 25: v = IIf(IIf(True, False, True), "a", "b")   ' nested
        Case 26: Set c = New Collection: Set o = IIf(True, c, Nothing): v = TypeName(o)    ' object branch
        Case 27: Set o = IIf(False, c, Nothing): v = TypeName(o)               ' Nothing branch
        Case 28: v = IIf(1 < 2, #1/1/2000#, #1/2/2000#)   ' Date branches
        Case 29: v = IIf(True, 1.5, 2)                 ' Double true branch
        Case 30: v = IIf(False, 1.5, 2)                ' Integer false branch
        Case 31: v = IIf(CByte(1), "a", "b")           ' Byte condition
        Case 32: v = IIf(Expression:=True, TruePart:="a", FalsePart:="b")   ' named arguments
    End Select
    T = "IIf." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "IIf." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
