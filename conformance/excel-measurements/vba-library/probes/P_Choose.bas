Attribute VB_Name = "P_Choose"
Option Explicit
' Choose (Interaction; every argument is evaluated, out-of-range index gives Null).
Private cnt As Long
Private Const N As Long = 33

Private Function Tick(ByVal x As Variant) As Variant
    cnt = cnt + 1
    Tick = x
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    Dim va As Variant, c As Collection, o As Object, z As Long
    On Error GoTo fail
    Select Case k
        Case 1: v = Choose(1, "a", "b", "c")          ' first
        Case 2: v = Choose(2, "a", "b", "c")          ' middle
        Case 3: v = Choose(3, "a", "b", "c")          ' last
        Case 4: v = Choose(0, "a", "b", "c")          ' below range: Null
        Case 5: v = Choose(4, "a", "b", "c")          ' above range: Null
        Case 6: v = Choose(-1, "a", "b", "c")         ' negative: Null
        Case 7: v = Choose(1.5, "a", "b", "c")        ' rounds to 2
        Case 8: v = Choose(2.5, "a", "b", "c")        ' banker's rounding to 2, or truncation
        Case 9: v = Choose(0.5, "a", "b", "c")        ' 0.5: rounds to 0 (Null) or truncates
        Case 10: v = Choose(0.51, "a", "b", "c")       ' rounds to 1
        Case 11: v = Choose(3.4, "a", "b", "c")        ' rounds to 3
        Case 12: v = Choose(3.6, "a", "b", "c")        ' rounds to 4: Null
        Case 13: v = Choose(Null, "a", "b")            ' Null index
        Case 14: v = Choose(Empty, "a", "b")           ' Empty index (0)
        Case 15: v = Choose("2", "a", "b")             ' numeric string index
        Case 16: v = Choose("x", "a", "b")             ' error 13
        Case 17: v = Choose(True, "a", "b")            ' -1: Null
        Case 18: v = Choose(1, Null, "b")              ' Null choice
        Case 19: v = Choose(1, 1, "a")                 ' result type Integer
        Case 20: v = Choose(2, 1, "a")                 ' result type String
        Case 21: v = Choose(1, 1.5, 2)                 ' no common type is forced
        Case 22: cnt = 0: v = Choose(1, Tick("a"), Tick("b"), Tick("c")) & "|" & cnt      ' all arguments evaluated
        Case 23: cnt = 0: v = Choose(9, Tick("a"), Tick("b")) & "|" & cnt                 ' evaluated though out of range
        Case 24: cnt = 0: v = Choose(2, 1, Tick(2), 1 \ z): v = cnt                       ' failing argument: error 11
        Case 25: v = Choose(2, 1, 1 \ z, 3)            ' a failing unselected choice still raises: error 11
        Case 26: v = Choose(1E10, "a")                 ' index beyond Long: error 6
        Case 27: v = Choose(2147483647, "a")           ' huge index: Null
        Case 28: v = Choose(1, Array(1, 2))            ' array result
        Case 29: Set c = New Collection: Set o = Choose(1, c): v = TypeName(o)    ' object choice
        Case 30: v = TypeName(Choose(1, 1, 2))
        Case 31: v = Choose(CByte(2), "a", "b")        ' Byte index
        Case 32: v = Choose(2, #1/1/2000#, #1/2/2000#) ' Date choices
        Case 33: v = Choose(1, CVErr(5), 2)            ' error value choice
    End Select
    T = "Choose." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Choose." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
