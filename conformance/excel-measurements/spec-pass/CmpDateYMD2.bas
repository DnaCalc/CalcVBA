Attribute VB_Name = "CmpDateYMD2"
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "dateYMD2=" & Sh(#99/12/31#) & vbLf: Exit Function
fail:
    Run = "dateYMD2=ERR" & Err.Number & vbLf
End Function
