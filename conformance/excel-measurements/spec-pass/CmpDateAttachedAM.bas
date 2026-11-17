Attribute VB_Name = "CmpDateAttachedAM"
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "dateAttachedAM=" & Sh(#8:30AM#) & vbLf: Exit Function
fail:
    Run = "dateAttachedAM=ERR" & Err.Number & vbLf
End Function
