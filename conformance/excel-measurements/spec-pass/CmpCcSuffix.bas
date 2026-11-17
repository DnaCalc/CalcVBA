Attribute VB_Name = "CmpCcSuffix"
#Const X% = 1
#If X Then
Private Const R = "yes"
#Else
Private Const R = "no"
#End If
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "ccSuffix=" & Sh(R) & vbLf: Exit Function
fail:
    Run = "ccSuffix=ERR" & Err.Number & vbLf
End Function
