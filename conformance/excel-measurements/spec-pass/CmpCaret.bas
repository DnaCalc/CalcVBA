Attribute VB_Name = "CmpCaret"
Dim x^
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    x = 5
    Run = "caret=" & Sh(x) & vbLf: Exit Function
fail:
    Run = "caret=ERR" & Err.Number & vbLf
End Function
