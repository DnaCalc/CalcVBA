Attribute VB_Name = "CmpNumD"
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "numD=" & Sh(1D3) & vbLf: Exit Function
fail:
    Run = "numD=ERR" & Err.Number & vbLf
End Function
