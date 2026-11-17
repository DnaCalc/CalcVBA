Attribute VB_Name = "CmpByRefVariantElem"
Private Sub IncL(x As Long)
    x = x + 1
End Sub
Private Function H() As Variant
    Dim vv As Variant
    vv = Array(1, 2)
    IncL vv(0)
    H = vv(0)
End Function
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "byRefVariantElem=" & Sh(H()) & vbLf: Exit Function
fail:
    Run = "byRefVariantElem=ERR" & Err.Number & vbLf
End Function
