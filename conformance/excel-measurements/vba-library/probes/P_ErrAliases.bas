Attribute VB_Name = "P_ErrAliases"
Option Explicit
' UNMEASURED alias/identity coverage beyond the eleven captured Err-value cases.
Private Const N As Long = 15
Private Function T(ByVal k As Long) As String
    Dim e As ErrObject, o As Object, boxed As Variant, v As Variant
    On Error GoTo fail
    Set e = Err
    Set o = Err
    Set boxed = Err
    Err.Clear
    Select Case k
        Case 1: v = (e Is o)
        Case 2: v = (ObjPtr(e) = ObjPtr(Err))
        Case 3: v = TypeName(e) & "," & TypeName(o) & "," & TypeName(boxed)
        Case 4: Err.Number = 23: v = e.Number & "," & o.Number & "," & boxed.Number
        Case 5: Err.Number = 23: v = e
        Case 6: Err.Number = 23: v = VarType(e) & "," & VarType(o) & "," & IsObject(e)
        Case 7: e.Description = "alias": v = Err.Description
        Case 8: e.Source = "AliasSource": v = Err.Source
        Case 9: e.HelpFile = "alias.hlp": e.HelpContext = 77: v = Err.HelpFile & "," & Err.HelpContext
        Case 10: e.Number = 23: e.Clear: v = Err.Number & "," & Err.Description
        Case 11: v = CallByName(e, "Number", VbGet)
        Case 12: CallByName e, "Description", VbLet, "late alias": v = Err.Description
        Case 13: CallByName e, "Raise", VbMethod, 5, "AliasSource", "Raised"
        Case 14: e.Raise Description:="named", Number:=11
        Case 15: v = TypeOf e Is ErrObject
    End Select
    T = "ErrAliases." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "ErrAliases." & k & "=ERR" & Err.Number & ":" & Err.Source & ":" & Err.Description
End Function
Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
