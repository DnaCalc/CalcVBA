Attribute VB_Name = "P_String"
Option Explicit
' String / String$ (Strings): repeat a character.
Private Const N As Long = 24

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = String(3, "ab")                   ' first character of a string
        Case 2: v = String(3, 65)                     ' character code
        Case 3: v = String(2, 256)                    ' code above 255
        Case 4: v = String(2, 321)                    ' code above 255 (mod 256?)
        Case 5: v = String(2, 65601)                  ' code above 65535
        Case 6: v = String(-1, "a")                   ' negative count: error 5
        Case 7: v = String(0, "a")                    ' zero count
        Case 8: v = String(2, "")                     ' empty character string: error 5
        Case 9: v = String(2, Null)                   ' Null character
        Case 10: v = String(Null, "a")                ' Null count: error 94
        Case 11: v = String(2, -1)                    ' negative code: error 5
        Case 12: v = String(2, "65")                  ' numeric string: first char "6"
        Case 13: v = String(2, ChrW(&H4E2D))          ' non-ASCII string character
        Case 14: v = String$(3, "x")                  ' $ form: String
        Case 15: v = String$(3, 66)                   ' $ form with a code
        Case 16: v = String(Number:=2, Character:="z")  ' named arguments
        Case 17: v = String(2, 65.5)                  ' fractional code
        Case 18: v = String(2, 0)                     ' code 0: NULs
        Case 19: v = String(2, 128)                   ' code 128 (ANSI mapping). locale
        Case 20: v = String(2.5, "a")                 ' fractional count rounds
        Case 21: v = String$(2, Null)                 ' $ form with Null: error 94
        Case 22: v = TypeName(String(2, "a"))         ' result type
        Case 23: v = String(2, Empty)                 ' Empty character
        Case 24: v = String("2", "a")                 ' numeric string count
    End Select
    T = "String." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "String." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
