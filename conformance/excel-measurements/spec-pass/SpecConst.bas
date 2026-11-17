Attribute VB_Name = "SpecConst"
' No Option Explicit: an undefined constant reads as an implicit Empty local instead of failing.
' MS-VBAL §6.1.1 enums and §6.1.2 constant modules that the code inspection found missing.

Private Function Sh(ByVal v As Variant) As String
    Sh = TypeName(v) & ":" & CStr(v)
End Function

Public Function Run() As String
    Dim s As String
    s = s & "vbKeyReturn=" & Sh(vbKeyReturn) & vbLf
    s = s & "vbKeyF1=" & Sh(vbKeyF1) & vbLf
    s = s & "vbButtonFace=" & Sh(vbButtonFace) & vbLf
    s = s & "vbWindowText=" & Sh(vbWindowText) & vbLf
    s = s & "vbHide=" & Sh(vbHide) & vbLf
    s = s & "vbNormalFocus=" & Sh(vbNormalFocus) & vbLf
    s = s & "vbMaximizedFocus=" & Sh(vbMaximizedFocus) & vbLf
    s = s & "vbCalGreg=" & Sh(vbCalGreg) & vbLf
    s = s & "vbCalHijri=" & Sh(vbCalHijri) & vbLf
    s = s & "vbFormControlMenu=" & Sh(vbFormControlMenu) & vbLf
    s = s & "vbAppTaskManager=" & Sh(vbAppTaskManager) & vbLf
    s = s & "vbIMEModeNoControl=" & Sh(vbIMEModeNoControl) & vbLf
    s = s & "vbIMEModeOn=" & Sh(vbIMEModeOn) & vbLf
    s = s & "vbAlias=" & Sh(vbAlias) & vbLf
    s = s & "vbNormalNoFocus=" & Sh(vbNormalNoFocus) & vbLf
    Run = s
End Function
