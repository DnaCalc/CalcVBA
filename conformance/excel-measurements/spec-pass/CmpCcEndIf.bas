Attribute VB_Name = "CmpCcEndIf"
#If True Then
Private Const R = "yes"
#EndIf
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "ccEndIf=" & Sh(R) & vbLf: Exit Function
fail:
    Run = "ccEndIf=ERR" & Err.Number & vbLf
End Function
