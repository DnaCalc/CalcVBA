Attribute VB_Name = "CmpNumBig"
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "numBig=" & Sh(99999999999999999999) & vbLf: Exit Function
fail:
    Run = "numBig=ERR" & Err.Number & vbLf
End Function
