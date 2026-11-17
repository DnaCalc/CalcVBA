Attribute VB_Name = "CmpCcString"
#Const S = "x"
#If S Then
Private Const R = "yes"
#Else
Private Const R = "no"
#End If
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "ccString=" & Sh(R) & vbLf: Exit Function
fail:
    Run = "ccString=ERR" & Err.Number & vbLf
End Function
