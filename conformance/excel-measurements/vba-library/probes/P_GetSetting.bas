Attribute VB_Name = "P_GetSetting"
Option Explicit
' GetSetting (Interaction; registry settings under AppName CalcVbaProbe, all deleted after each case).
Private Const A As String = "CalcVbaProbe"
Private Const N As Long = 20

Private Function T(ByVal k As Long) As String
    Dim v As Variant, en As Long
    On Error GoTo fail
    Select Case k
        Case 1: v = GetSetting(A, "S", "K")                                               ' missing, no default: ""
        Case 2: v = GetSetting(A, "S", "K", "dflt")                                       ' missing, default returned
        Case 3: v = GetSetting(A, "S", "K", 42)                                           ' numeric default returned as String
        Case 4: v = GetSetting(A, "S", "K", 3.5)                                          ' Double default
        Case 5: v = GetSetting(A, "S", "K", Null)                                         ' Null default
        Case 6: v = GetSetting(A, "S", "K", Empty)                                        ' Empty default
        Case 7: v = GetSetting(A, "S", "K", "")                                           ' empty default
        Case 8: SaveSetting A, "S", "K", "stored": v = GetSetting(A, "S", "K", "dflt")    ' existing value ignores default
        Case 9: SaveSetting A, "S", "K", "": v = GetSetting(A, "S", "K", "dflt")          ' stored empty value
        Case 10: SaveSetting A, "S", "K", "x": v = TypeName(GetSetting(A, "S", "K"))       ' result type String
        Case 11: SaveSetting A, "S", "K", "x": v = GetSetting(AppName:=A, Section:="S", Key:="K", Default:="d")   ' named arguments
        Case 12: SaveSetting A, "S", "K", "x": v = GetSetting(A, "S2", "K", "other section")   ' missing section
        Case 13: v = GetSetting("CalcVbaProbeNoSuchApp", "S", "K", "d")                    ' missing application
        Case 14: SaveSetting A, "S", "K", "x": v = GetSetting(A, "s", "k", "case")         ' case-insensitive section and key
        Case 15: SaveSetting A, "S", "K", "  sp  ": v = GetSetting(A, "S", "K")            ' spaces preserved
        Case 16: v = GetSetting(Null, "S", "K")                                            ' Null application: error 94
        Case 17: v = GetSetting(A, Null, "K")                                              ' Null section: error 94
        Case 18: v = GetSetting(A, "S", 1, "d")                                            ' numeric key converted to text
        Case 19: SaveSetting A, "S", "1", "num": v = GetSetting(A, "S", 1)                 ' numeric key matches text key
        Case 20: v = GetSetting(A, "S", "K", True)                                         ' Boolean default
    End Select
    On Error Resume Next: DeleteSetting A
    T = "GetSetting." & k & "=" & Sh(v)
    Exit Function
fail:
    en = Err.Number
    Resume cleanup
cleanup:
    On Error Resume Next: DeleteSetting A
    T = "GetSetting." & k & "=ERR" & en
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
    On Error Resume Next: DeleteSetting A
End Function
