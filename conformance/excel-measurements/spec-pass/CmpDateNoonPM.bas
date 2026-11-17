Attribute VB_Name = "CmpDateNoonPM"
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "dateNoonPM=" & Sh(#1/1/2026 0:00 PM#) & vbLf: Exit Function
fail:
    Run = "dateNoonPM=ERR" & Err.Number & vbLf
End Function
