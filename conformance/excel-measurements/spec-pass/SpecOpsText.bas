Attribute VB_Name = "SpecOpsText"
Option Explicit
Option Compare Text
' MS-VBAL §5.6.9.5 / §5.6.9.6 under Option Compare Text (non-ASCII folding, ranges, ligatures).

Private Function T(ByVal k As String) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case "likeUmlaut": v = (ChrW(196) Like ChrW(228))
        Case "likeUmlautRange": v = (ChrW(228) Like "[a-b]")
        Case "likeUmlautRangeZ": v = (ChrW(228) Like "[z-" & ChrW(255) & "]")
        Case "likeLigature": v = (ChrW(230) Like "ae")
        Case "eqUmlaut": v = (ChrW(196) = ChrW(228))
        Case "eqLigature": v = (ChrW(230) = "ae")
        Case "likeUpperRange": v = ("a" Like "[A-C]")
        Case "likeQSurrogate": v = (ChrW(&HD83D) & ChrW(&HDE00) Like "??")
        Case Else: v = "unknown case"
    End Select
    T = k & "=" & TypeName(v) & ":" & CStr(v): Exit Function
fail:
    T = k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim keys As Variant, i As Long, s As String
    keys = Split("likeUmlaut likeUmlautRange likeUmlautRangeZ likeLigature eqUmlaut eqLigature likeUpperRange likeQSurrogate", " ")
    For i = 0 To UBound(keys)
        s = s & T(keys(i)) & vbLf
    Next
    Run = s
End Function
