Attribute VB_Name = "P_DeleteSetting"
Option Explicit
' DeleteSetting (Interaction; registry settings under AppName CalcVbaProbe, all deleted after each case).
Private Const A As String = "CalcVbaProbe"
Private Const N As Long = 20

Private Function T(ByVal k As Long) As String
    Dim v As Variant, en As Long
    Dim g As Variant
    On Error GoTo fail
    Select Case k
        Case 1: SaveSetting A, "S", "K", "1": DeleteSetting A, "S", "K": v = GetSetting(A, "S", "K", "gone")   ' delete a key
        Case 2: SaveSetting A, "S", "K", "1": DeleteSetting A, "S": v = GetSetting(A, "S", "K", "gone")        ' delete a section
        Case 3: SaveSetting A, "S", "K", "1": DeleteSetting A: v = GetSetting(A, "S", "K", "gone")             ' delete the application
        Case 4: SaveSetting A, "S", "K", "1": DeleteSetting A, "S", "NoKey"               ' missing key: error 5
        Case 5: SaveSetting A, "S", "K", "1": DeleteSetting A, "NoSection"                ' missing section: error 5
        Case 6: DeleteSetting "CalcVbaProbeNoSuchApp"                                     ' missing application: error 5
        Case 7: DeleteSetting A                                                           ' application not saved: error 5
        Case 8: SaveSetting A, "S", "K", "1": DeleteSetting A, "S", "K": DeleteSetting A, "S", "K"   ' second delete: error 5
        Case 9: SaveSetting A, "S", "K1", "1": SaveSetting A, "S", "K2", "2": DeleteSetting A, "S", "K1": v = GetSetting(A, "S", "K2", "gone")   ' other key kept
        Case 10: SaveSetting A, "S1", "K", "1": SaveSetting A, "S2", "K", "2": DeleteSetting A, "S1": v = GetSetting(A, "S2", "K", "gone")        ' other section kept
        Case 11: SaveSetting A, "S1", "K", "1": SaveSetting A, "S2", "K", "2": DeleteSetting A: v = GetSetting(A, "S2", "K", "gone")             ' whole application removed
        Case 12: SaveSetting A, "S", "K", "1": DeleteSetting AppName:=A, Section:="S", Key:="K": v = GetSetting(A, "S", "K", "gone")   ' named arguments
        Case 13: SaveSetting A, "S", "K", "1": DeleteSetting A, "s", "k": v = GetSetting(A, "S", "K", "gone")   ' case-insensitive
        Case 14: SaveSetting A, "S", "K", "1": DeleteSetting Null                          ' Null application: error 94
        Case 15: SaveSetting A, "S", "K", "1": DeleteSetting A, , "K"                      ' Key without Section: error 5
        Case 16: SaveSetting A, "S", "K", "1": DeleteSetting A, "S", "K": g = GetAllSettings(A, "S"): v = TypeName(g) & "|" & IsArray(g)   ' section remains but empty
        Case 17: SaveSetting A, "S", "K", "1": DeleteSetting A, "S": g = GetAllSettings(A, "S"): v = TypeName(g) & "|" & IsArray(g)         ' section deleted
        Case 18: SaveSetting A, "S", "K", "1": DeleteSetting A, "", "K"                    ' empty section name with a key
        Case 19: SaveSetting A, "S", "K", "1": DeleteSetting A, "S", ""                    ' empty key name
        Case 20: SaveSetting A, "S", "K", "1": DeleteSetting A, "S", Empty                 ' Empty key
    End Select
    On Error Resume Next: DeleteSetting A
    T = "DeleteSetting." & k & "=" & Sh(v)
    Exit Function
fail:
    en = Err.Number
    Resume cleanup
cleanup:
    On Error Resume Next: DeleteSetting A
    T = "DeleteSetting." & k & "=ERR" & en
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
    On Error Resume Next: DeleteSetting A
End Function
