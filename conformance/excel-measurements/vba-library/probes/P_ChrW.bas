Attribute VB_Name = "P_ChrW"
Option Explicit
' ChrW / ChrW$ (Strings). Code units -32768..65535; Sh escapes non-ASCII as {code}.
Private Const N As Long = 22

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = ChrW(65)                         ' typical
        Case 2: v = ChrW(233)                        ' e-acute
        Case 3: v = ChrW(223)                        ' sharp s
        Case 4: v = ChrW(8364)                       ' euro
        Case 5: v = ChrW(26085)                      ' CJK
        Case 6: v = ChrW(55357) & ChrW(56832)        ' surrogate pair (U+1F600)
        Case 7: v = ChrW(0)                          ' NUL
        Case 8: v = ChrW(-32768)                     ' smallest: &H8000
        Case 9: v = ChrW(-1)                         ' &HFFFF
        Case 10: v = ChrW(65535)                     ' largest
        Case 11: v = ChrW(65536)                     ' error 5
        Case 12: v = ChrW(-32769)                    ' error 5
        Case 13: v = ChrW(Null)                      ' Null: error 94
        Case 14: v = ChrW("233")                     ' numeric string
        Case 15: v = ChrW$(65)                       ' $ form: String
        Case 16: v = ChrW(CharCode:=66)              ' named argument
        Case 17: v = ChrW(160)                       ' no-break space
        Case 18: v = ChrW(128)                       ' C1 control, not the euro
        Case 19: v = ChrW(Empty)                     ' Empty is 0
        Case 20: v = ChrW(65.5)                      ' rounds to even: 66
        Case 21: v = ChrW("x")                       ' non-numeric string: error 13
        Case 22: v = ChrW$(CharCode:=8364)           ' $ form, named
    End Select
    T = "ChrW." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "ChrW." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
