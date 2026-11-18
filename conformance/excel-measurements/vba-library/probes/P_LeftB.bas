Attribute VB_Name = "P_LeftB"
Option Explicit
' LeftB / LeftB$ (Strings). Length counts bytes; results are shown through Bytes (type, LenB, hex bytes).
Private Const N As Long = 18

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
        Case 1: x = LeftB("abcdef", 4): v = Bytes(x)         ' typical: "ab" is 4 bytes
        Case 2: x = LeftB("abc", 3): v = Bytes(x)            ' odd byte count
        Case 3: x = LeftB("abc", 0): v = Bytes(x)            ' zero length
        Case 4: x = LeftB("abc", 100): v = Bytes(x)          ' past the end
        Case 5: x = LeftB("abc", -1): v = Bytes(x)           ' negative: error 5
        Case 6: x = LeftB(Null, 2): v = Bytes(x)             ' Null propagates
        Case 7: x = LeftB("abc", Null): v = Bytes(x)         ' Null length: error 94
        Case 8: x = LeftB$("abc", 2): v = Bytes(x)           ' $ form
        Case 9: x = LeftB(String:="abc", Length:=2): v = Bytes(x)   ' named arguments
        Case 10: x = LeftB(12345, 4): v = Bytes(x)           ' number as string
        Case 11: x = LeftB(ChrW(8364), 1): v = Bytes(x)      ' low byte of the euro
        Case 12: x = LeftB("", 2): v = Bytes(x)              ' empty
        Case 13: x = LeftB(Empty, 2): v = Bytes(x)           ' Empty
        Case 14: x = LeftB("abc", "2"): v = Bytes(x)         ' numeric string length
        Case 15: x = LeftB("abc", 1.5): v = Bytes(x)         ' rounds to even: 2
        Case 16: x = LeftB$(Null, 2): v = Bytes(x)           ' $ form with Null: error 94
        Case 17: x = LeftB(ChrW(55357) & ChrW(56832), 3): v = Bytes(x)   ' into a surrogate pair
        Case 18: x = LeftB$(String:="abc", Length:=2): v = Bytes(x)      ' $ form, named
    End Select
    T = "LeftB." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "LeftB." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
