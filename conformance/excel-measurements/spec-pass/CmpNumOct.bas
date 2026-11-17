Attribute VB_Name = "CmpNumOct"
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "numOct=" & Sh(&17) & vbLf: Exit Function
fail:
    Run = "numOct=ERR" & Err.Number & vbLf
End Function
