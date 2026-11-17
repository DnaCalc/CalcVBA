Attribute VB_Name = "CmpDateMonYear"
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "dateMonYear=" & Sh(#3/2020#) & vbLf: Exit Function
fail:
    Run = "dateMonYear=ERR" & Err.Number & vbLf
End Function
