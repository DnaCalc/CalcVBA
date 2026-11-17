Attribute VB_Name = "CmpOptBase1Amp"
Option Base 1&
Dim a(3)
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "optBase1Amp=" & Sh(LBound(a)) & vbLf: Exit Function
fail:
    Run = "optBase1Amp=ERR" & Err.Number & vbLf
End Function
