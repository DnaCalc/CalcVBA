Attribute VB_Name = "CmpCatBytes"
Dim bb() As Byte
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    bb = "AB"
    Run = "catBytes=" & Sh(bb & "!") & vbLf: Exit Function
fail:
    Run = "catBytes=ERR" & Err.Number & vbLf
End Function
