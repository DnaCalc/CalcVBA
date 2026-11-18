Attribute VB_Name = "P_AscB"
Option Explicit
' AscB (Strings). Returns Byte: the first byte of the UTF-16LE string.
Private Const N As Long = 14

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = AscB("A")                        ' typical
        Case 2: v = AscB("AB")                       ' first byte only
        Case 3: v = AscB("")                         ' empty: error 5
        Case 4: v = AscB(Null)                       ' Null: error 94
        Case 5: v = AscB(ChrW(233))                  ' low byte of &HE9
        Case 6: v = AscB(ChrW(8364))                 ' low byte of &H20AC = 172
        Case 7: v = AscB(ChrW(256))                  ' low byte 0
        Case 8: v = AscB(ChrW(26085))                ' low byte of &H65E5 = 229
        Case 9: v = AscB(123)                        ' number as string
        Case 10: v = AscB(Empty)                     ' Empty: error 5
        Case 11: v = AscB(String:="Z")               ' named argument
        Case 12: v = AscB(ChrB(200))                 ' single-byte string
        Case 13: v = AscB(ChrW(&HFF00))              ' low byte 0, high 255
        Case 14: v = AscB(Chr(128))                  ' euro via 1252 -> &H20AC ' locale
    End Select
    T = "AscB." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "AscB." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
