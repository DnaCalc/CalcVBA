Attribute VB_Name = "CmpCcDup"
#If Win64 Then
#Const P = 8
#Else
#Const P = 4
#End If
#If P = 8 Then
Private Const R = "8"
#Else
Private Const R = "4"
#End If
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "ccDup=" & Sh(R) & vbLf: Exit Function
fail:
    Run = "ccDup=ERR" & Err.Number & vbLf
End Function
