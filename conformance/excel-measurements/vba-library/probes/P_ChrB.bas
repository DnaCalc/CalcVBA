Attribute VB_Name = "P_ChrB"
Option Explicit
' ChrB / ChrB$ (Strings). Yields a one-byte string, shown through Bytes (type, LenB, hex bytes).
Private Const N As Long = 15

Private Function Bytes(x As Variant) As String
    Dim i As Long, r As String
    If IsNull(x) Then Bytes = "Null": Exit Function
    r = TypeName(x) & " lenb=" & LenB(x) & " ["
    For i = 1 To LenB(x)
        r = r & Right$("0" & Hex$(AscB(MidB(x, i, 1))), 2)
    Next
    Bytes = r & "]"
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: x = ChrB(65): v = Bytes(x)           ' typical
        Case 2: x = ChrB(0): v = Bytes(x)            ' zero byte
        Case 3: x = ChrB(255): v = Bytes(x)          ' largest byte
        Case 4: x = ChrB(256): v = Bytes(x)          ' out of range
        Case 5: x = ChrB(-1): v = Bytes(x)           ' negative
        Case 6: x = ChrB(Null): v = Bytes(x)         ' Null: error 94
        Case 7: x = ChrB("66"): v = Bytes(x)         ' numeric string
        Case 8: x = ChrB$(67): v = Bytes(x)          ' $ form
        Case 9: x = ChrB(CharCode:=68): v = Bytes(x) ' named argument
        Case 10: x = ChrB(65.5): v = Bytes(x)        ' rounds to even: 66
        Case 11: x = ChrB(Empty): v = Bytes(x)       ' Empty is 0
        Case 12: x = ChrB(65) & ChrB(0): v = Bytes(x)    ' two bytes joined
        Case 13: v = (ChrB(65) & ChrB(0)) = "A"      ' same bytes as "A"
        Case 14: v = Len(ChrB(65))                   ' half a character
        Case 15: x = ChrB(233) & ChrB(1): v = Bytes(x)   ' bytes of ChrW(&H1E9)
    End Select
    T = "ChrB." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "ChrB." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
