Attribute VB_Name = "CmpCcVba6"
#If VBA6 Then
Private Const R = "vba6"
#Else
Private Const R = "not"
#End If
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "ccVba6=" & Sh(R) & vbLf: Exit Function
fail:
    Run = "ccVba6=ERR" & Err.Number & vbLf
End Function
