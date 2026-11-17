Attribute VB_Name = "CmpCcCont"
#Const A = 1 _
    + 1
#If A = 2 Then
Private Const R = "two"
#Else
Private Const R = "other"
#End If
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "ccCont=" & Sh(R) & vbLf: Exit Function
fail:
    Run = "ccCont=ERR" & Err.Number & vbLf
End Function
