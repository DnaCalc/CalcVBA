Attribute VB_Name = "P_GetAllSettings"
Option Explicit
' GetAllSettings (Interaction; 2-D Variant array of key/value pairs, registry under CalcVbaProbe).
Private Const A As String = "CalcVbaProbe"
Private Const N As Long = 19

Private Function Dump(g As Variant) As String
    Dim i As Long, r As String
    If Not IsArray(g) Then
        Dump = "notarray:" & TypeName(g)
        Exit Function
    End If
    r = TypeName(g) & "|" & LBound(g, 1) & ":" & UBound(g, 1) & "|" & LBound(g, 2) & ":" & UBound(g, 2) & "|"
    For i = LBound(g, 1) To UBound(g, 1)
        r = r & Esc(CStr(g(i, 0))) & "=" & Esc(CStr(g(i, 1))) & ";"
    Next
    Dump = r
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant, en As Long
    Dim g As Variant
    On Error GoTo fail
    Select Case k
        Case 1: SaveSetting A, "S", "K1", "a": SaveSetting A, "S", "K2", "b": SaveSetting A, "S", "K3", "c": g = GetAllSettings(A, "S"): v = Dump(g)   ' three keys
        Case 2: SaveSetting A, "S", "K1", "a": g = GetAllSettings(A, "S"): v = Dump(g)    ' single key
        Case 3: g = GetAllSettings(A, "S"): v = Dump(g)                                   ' missing section
        Case 4: g = GetAllSettings("CalcVbaProbeNoSuchApp", "S"): v = Dump(g)             ' missing application
        Case 5: SaveSetting A, "S", "b", "2": SaveSetting A, "S", "a", "1": SaveSetting A, "S", "c", "3": g = GetAllSettings(A, "S"): v = Dump(g)   ' order of keys
        Case 6: SaveSetting A, "S", "K1", "a": SaveSetting A, "S", "K2", "b": DeleteSetting A, "S", "K1": g = GetAllSettings(A, "S"): v = Dump(g)   ' after a key is deleted
        Case 7: SaveSetting A, "S", "K", "a=b;c": g = GetAllSettings(A, "S"): v = Dump(g) ' special characters
        Case 8: SaveSetting A, "S", "K", "": g = GetAllSettings(A, "S"): v = Dump(g)      ' empty value
        Case 9: SaveSetting A, "S", "K", "1": g = GetAllSettings(A, "S"): v = TypeName(g(0, 1))   ' element type String
        Case 10: SaveSetting A, "S", "K", "1": g = GetAllSettings(A, "S"): v = TypeName(g)         ' array type
        Case 11: SaveSetting A, "S", "K", "1": g = GetAllSettings(A, "S"): v = UBound(g, 1) & "," & UBound(g, 2)   ' bounds
        Case 12: SaveSetting A, "S", "K", "1": g = GetAllSettings(AppName:=A, Section:="S"): v = Dump(g)   ' named arguments
        Case 13: SaveSetting A, "S", "K", "1": g = GetAllSettings(A, ""): v = Dump(g)      ' empty section name
        Case 14: SaveSetting A, "S", "K", "1": g = GetAllSettings(A, "s"): v = Dump(g)     ' section case-insensitive
        Case 15: SaveSetting A, "S1", "K", "1": SaveSetting A, "S2", "K", "2": g = GetAllSettings(A, "S2"): v = Dump(g)   ' only the named section
        Case 16: g = GetAllSettings(Null, "S"): v = Dump(g)                                ' Null application: error 94
        Case 17: SaveSetting A, "S", "K", "1": g = GetAllSettings(A, "S"): v = g(0, 0) & "=" & g(0, 1)   ' element access
        Case 18: SaveSetting A, "S", "K", "1": g = GetAllSettings(A, "S"): v = g(1, 0)     ' out of range: error 9
        Case 19: SaveSetting A, "S", "K", "1": v = TypeName(GetAllSettings(A, "S"))        ' direct result type
    End Select
    On Error Resume Next: DeleteSetting A
    T = "GetAllSettings." & k & "=" & Sh(v)
    Exit Function
fail:
    en = Err.Number
    Resume cleanup
cleanup:
    On Error Resume Next: DeleteSetting A
    T = "GetAllSettings." & k & "=ERR" & en
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
    On Error Resume Next: DeleteSetting A
End Function
