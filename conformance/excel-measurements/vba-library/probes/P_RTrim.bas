Attribute VB_Name = "P_RTrim"
Option Explicit
' RTrim / RTrim$ (Strings): removes trailing U+0020 only.
Private Const N As Long = 16

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = RTrim("  abc  ")                  ' typical: leading spaces kept
        Case 2: v = RTrim("abc")                      ' nothing to trim
        Case 3: v = RTrim("   ")                      ' all spaces
        Case 4: v = RTrim("")                         ' empty
        Case 5: v = RTrim(Null)                       ' Null propagates
        Case 6: v = RTrim(Empty)                      ' Empty
        Case 7: v = RTrim(123)                        ' number as string argument
        Case 8: v = RTrim$("  abc  ")                 ' $ form: String
        Case 9: v = RTrim$(Null)                      ' $ form with Null: error 94
        Case 10: v = RTrim("abc" & vbTab)             ' tab is not trimmed
        Case 11: v = RTrim("abc" & ChrW(160))         ' non-breaking space is not trimmed
        Case 12: v = RTrim("abc" & ChrW(12288))       ' ideographic space is not trimmed
        Case 13: v = RTrim("abc" & vbCrLf & " ")      ' CR LF not trimmed
        Case 14: v = RTrim(String:="abc  ")           ' named argument
        Case 15: v = RTrim("a  b  ")                  ' interior spaces kept
        Case 16: v = RTrim(1.5 & " ")                 ' number concatenated
    End Select
    T = "RTrim." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "RTrim." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
