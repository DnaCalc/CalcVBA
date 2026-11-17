Attribute VB_Name = "CmpCallByVal"
Private Sub G(x)
    x = 1
End Sub
Private Function H() As Variant
    Dim k
    G ByVal k
    H = IsEmpty(k)
End Function
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "callByVal=" & Sh(H()) & vbLf: Exit Function
fail:
    Run = "callByVal=ERR" & Err.Number & vbLf
End Function
