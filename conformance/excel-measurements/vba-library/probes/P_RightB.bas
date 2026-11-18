Attribute VB_Name = "P_RightB"
Option Explicit
' RightB / RightB$ (Strings; byte-based).
Private Const N As Long = 13

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = RightB("abcdef", 4)               ' even bytes: "ef"
        Case 2: v = RightB("abcdef", 3)               ' odd bytes: splits a character
        Case 3: v = RightB("abc", 0)                  ' zero length
        Case 4: v = RightB("abc", 100)                ' past the start
        Case 5: v = RightB("abc", -1)                 ' negative length: error 5
        Case 6: v = RightB(Null, 2)                   ' Null propagates
        Case 7: v = RightB("abc", Null)               ' Null length: error 94
        Case 8: v = RightB$("abcdef", 2)              ' $ form: String
        Case 9: v = RightB$(Null, 2)                  ' $ form with Null: error 94
        Case 10: v = RightB(12345, 4)                 ' number as string argument
        Case 11: v = LenB(RightB("abcdef", 3))        ' byte length of the odd slice
        Case 12: v = RightB("", 2)                    ' empty string
        Case 13: v = RightB(String:="abc", Length:=2) ' named arguments
    End Select
    T = "RightB." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "RightB." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
