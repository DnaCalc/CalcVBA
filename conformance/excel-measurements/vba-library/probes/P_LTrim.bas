Attribute VB_Name = "P_LTrim"
Option Explicit
' LTrim / LTrim$ (Strings). Only the ASCII space (32) is trimmed.
Private Const N As Long = 20

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = LTrim("  abc  ")                 ' typical: trailing spaces kept
        Case 2: v = LTrim("")                        ' empty
        Case 3: v = LTrim(Null)                      ' Null propagates
        Case 4: v = LTrim(123)                       ' number: "123"
        Case 5: v = LTrim(Empty)                     ' Empty
        Case 6: v = LTrim(vbTab & "a")               ' tab not trimmed
        Case 7: v = LTrim(ChrW(160) & "a")           ' no-break space not trimmed
        Case 8: v = LTrim(vbCr & vbLf & " a")        ' CR and LF not trimmed
        Case 9: v = LTrim("   ")                     ' all spaces
        Case 10: v = LTrim$("  x")                   ' $ form: String
        Case 11: v = LTrim$(Null)                    ' $ form with Null: error 94
        Case 12: v = LTrim(String:="  y")            ' named argument
        Case 13: v = LTrim(ChrW(12288) & "a")        ' ideographic space not trimmed
        Case 14: v = LTrim(ChrW(8195) & "a")         ' em space not trimmed
        Case 15: v = LTrim(-5)                       ' negative number
        Case 16: v = LTrim(True)                     ' Boolean: "True" ' locale
        Case 17: v = LTrim(" " & ChrW(160) & " a")   ' stops at the no-break space
        Case 18: v = LTrim(ChrW(0) & " a")           ' NUL not trimmed
        Case 19: v = LTrim$(String:="  z")           ' $ form, named
        Case 20: v = LTrim("  a  b ")                ' inner spaces kept
    End Select
    T = "LTrim." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "LTrim." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
