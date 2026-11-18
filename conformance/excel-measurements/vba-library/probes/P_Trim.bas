Attribute VB_Name = "P_Trim"
Option Explicit
' Trim / Trim$ (Strings): removes U+0020 from both ends only.
Private Const N As Long = 16

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = Trim("  abc  ")                   ' typical
        Case 2: v = Trim(" a  b ")                    ' interior spaces kept
        Case 3: v = Trim("   ")                       ' all spaces
        Case 4: v = Trim("")                          ' empty
        Case 5: v = Trim(Null)                        ' Null propagates
        Case 6: v = Trim(Empty)                       ' Empty
        Case 7: v = Trim(123)                         ' number as string argument
        Case 8: v = Trim$("  abc  ")                  ' $ form: String
        Case 9: v = Trim$(Null)                       ' $ form with Null: error 94
        Case 10: v = Trim(vbTab & "abc" & vbTab)      ' tabs not trimmed
        Case 11: v = Trim(ChrW(160) & "abc" & ChrW(160))  ' non-breaking spaces not trimmed
        Case 12: v = Trim(ChrW(12288) & "abc")        ' ideographic space not trimmed
        Case 13: v = Trim(" " & ChrW(160) & " abc")   ' mixed space and nbsp
        Case 14: v = Trim(String:="  abc  ")          ' named argument
        Case 15: v = TypeName(Trim(" a "))            ' result type
        Case 16: v = TypeName(Trim$(" a "))           ' $ result type
    End Select
    T = "Trim." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Trim." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
