Attribute VB_Name = "CmpOptExplicit2"
Option Explicit
Option Explicit
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "optExplicit2=" & Sh(1) & vbLf: Exit Function
fail:
    Run = "optExplicit2=ERR" & Err.Number & vbLf
End Function
