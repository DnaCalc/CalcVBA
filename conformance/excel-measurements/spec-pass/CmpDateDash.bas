Attribute VB_Name = "CmpDateDash"
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "dateDash=" & Sh(#2-3-2026#) & vbLf: Exit Function
fail:
    Run = "dateDash=ERR" & Err.Number & vbLf
End Function
