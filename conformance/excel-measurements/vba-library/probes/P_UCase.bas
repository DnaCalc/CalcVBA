Attribute VB_Name = "P_UCase"
Option Explicit
' UCase / UCase$ (Strings).
Private Const N As Long = 20

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = UCase("abc def")                  ' typical
        Case 2: v = UCase("")                         ' empty
        Case 3: v = UCase(Null)                       ' Null propagates
        Case 4: v = UCase(Empty)                      ' Empty
        Case 5: v = UCase(12)                         ' number as string argument
        Case 6: v = UCase(ChrW(228) & ChrW(246) & ChrW(252))  ' Latin-1 letters
        Case 7: v = UCase(ChrW(223))                  ' sharp s stays? locale
        Case 8: v = UCase(ChrW(255))                  ' y-diaeresis to U+0178. locale
        Case 9: v = UCase("i")                        ' plain i (not Turkish)
        Case 10: v = UCase(ChrW(305))                 ' dotless i to I
        Case 11: v = UCase(ChrW(454))                 ' dz digraph U+01C6
        Case 12: v = UCase(ChrW(&H3C3) & ChrW(&H3C2)) ' Greek sigma and final sigma
        Case 13: v = UCase(ChrW(160) & "a")           ' non-breaking space unchanged
        Case 14: v = UCase$("abc")                    ' $ form: String
        Case 15: v = UCase$(Null)                     ' $ form with Null: error 94
        Case 16: v = UCase(String:="abc")             ' named argument
        Case 17: v = UCase(ChrW(&H430) & ChrW(&H431)) ' Cyrillic
        Case 18: v = UCase(ChrW(&HFB01))              ' fi ligature U+FB01
        Case 19: v = TypeName(UCase("a"))             ' result type
        Case 20: v = TypeName(UCase$("a"))            ' $ result type
    End Select
    T = "UCase." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "UCase." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
