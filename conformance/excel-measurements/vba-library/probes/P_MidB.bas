Attribute VB_Name = "P_MidB"
Option Explicit
' MidB / MidB$ (Strings; byte-based; MS-VBAL 6.1.2.11.1.x).
Private Const N As Long = 17

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = MidB("abcdef", 3, 4)              ' even byte positions: "bc"
        Case 2: v = MidB("abcdef", 2, 2)              ' odd start splits characters
        Case 3: v = MidB("abcdef", 2)                 ' odd start, length omitted
        Case 4: v = MidB("abc", 7)                    ' start = LenB+1: empty
        Case 5: v = MidB("abc", 0)                    ' start 0: error 5
        Case 6: v = MidB("abc", 1, 1)                 ' odd length: half a character
        Case 7: v = MidB("abc", 2, -1)                ' negative length: error 5
        Case 8: v = MidB("abc", 5, 100)               ' length past the end
        Case 9: v = MidB(Null, 1, 2)                  ' Null propagates
        Case 10: v = MidB("abc", Null, 2)             ' Null start: error 94
        Case 11: v = MidB(12345, 3, 4)                ' number as string argument
        Case 12: v = MidB$("abcdef", 3, 4)            ' $ form: String
        Case 13: v = MidB$(Null, 1)                   ' $ form with Null: error 94
        Case 14: v = LenB(MidB("abcdef", 2, 3))       ' byte length of an odd slice
        Case 15: v = MidB("", 1, 1)                   ' empty string
        Case 16: v = MidB(String:="abcdef", Start:=3, Length:=2)  ' named arguments
        Case 17: v = MidB("abc", 1, 0)                ' zero length
    End Select
    T = "MidB." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "MidB." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
