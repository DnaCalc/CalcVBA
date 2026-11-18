Attribute VB_Name = "P_Space"
Option Explicit
' Space / Space$ (Strings).
Private Const N As Long = 15

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = Space(3)                          ' typical
        Case 2: v = Space(0)                          ' zero
        Case 3: v = Space(-1)                         ' negative: error 5
        Case 4: v = Space(Null)                       ' Null: error 94
        Case 5: v = Space("2")                        ' numeric string
        Case 6: v = Space(2.5)                        ' fractional rounds (banker's)
        Case 7: v = Space(3.5)                        ' fractional rounds to 4
        Case 8: v = Space$(3)                         ' $ form: String
        Case 9: v = Space(Empty)                      ' Empty = 0
        Case 10: v = Space(Number:=2)                 ' named argument
        Case 11: v = Len(Space(1000))                 ' long result
        Case 12: v = Space("abc")                     ' non-numeric string: error 13
        Case 13: v = TypeName(Space(2))               ' result type
        Case 14: v = Space$(Number:=1)                ' named, $ form
        Case 15: v = Space$(Null)                     ' $ form with Null: error 94
    End Select
    T = "Space." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Space." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
