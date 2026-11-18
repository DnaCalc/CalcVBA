Attribute VB_Name = "P_SaveSetting"
Option Explicit
' SaveSetting (Interaction; registry settings under AppName CalcVbaProbe, all deleted after each case).
Private Const A As String = "CalcVbaProbe"
Private Const N As Long = 20

Private Function T(ByVal k As Long) As String
    Dim v As Variant, en As Long
    On Error GoTo fail
    Select Case k
        Case 1: SaveSetting A, "S", "K", "v1": v = GetSetting(A, "S", "K")                ' round trip
        Case 2: SaveSetting A, "S", "K", "v1": SaveSetting A, "S", "K", "v2": v = GetSetting(A, "S", "K")   ' overwrite
        Case 3: SaveSetting A, "S", "K", "a=b;c d": v = GetSetting(A, "S", "K")           ' special characters
        Case 4: SaveSetting A, "S", "K", "": v = GetSetting(A, "S", "K", "dflt")          ' empty value kept or default
        Case 5: SaveSetting A, "S", "K", ChrW(233) & ChrW(955): v = GetSetting(A, "S", "K")   ' Unicode value
        Case 6: SaveSetting A, "S", "K", String(300, "x"): v = Len(GetSetting(A, "S", "K"))   ' long value
        Case 7: SaveSetting A, "S", "Key", "v": v = GetSetting(A, "S", "key", "none")     ' key is case-insensitive
        Case 8: SaveSetting A, "S", "K1", "a": SaveSetting A, "S", "K2", "b": v = GetSetting(A, "S", "K1") & GetSetting(A, "S", "K2")   ' two keys
        Case 9: SaveSetting AppName:=A, Section:="S", Key:="K", Setting:="named": v = GetSetting(A, "S", "K")   ' named arguments
        Case 10: SaveSetting A, "My Section", "My Key", "v": v = GetSetting(A, "My Section", "My Key")   ' names with spaces
        Case 11: SaveSetting A, "S", "K", 42: v = GetSetting(A, "S", "K")                  ' number saved as text
        Case 12: SaveSetting A, "S", "K", Null                                             ' Null setting: error 94
        Case 13: SaveSetting A, "S", "K", Empty: v = GetSetting(A, "S", "K", "dflt")       ' Empty setting
        Case 14: SaveSetting LCase(A), "S", "K", "v": v = GetSetting(A, "S", "K")          ' AppName case-insensitive
        Case 15: SaveSetting A, "", "K", "v": v = GetSetting(A, "", "K", "dflt")           ' empty section
        Case 16: SaveSetting A, "S", "", "v": v = GetSetting(A, "S", "", "dflt")           ' empty key
        Case 17: SaveSetting A, "S", "K", "  spaced  ": v = GetSetting(A, "S", "K")        ' leading and trailing spaces
        Case 18: SaveSetting A, "S", "K", "line1" & vbCrLf & "line2": v = GetSetting(A, "S", "K")   ' embedded newline
        Case 19: SaveSetting A, "S", "K", 1.5: v = GetSetting(A, "S", "K")                 ' Double saved as text
        Case 20: SaveSetting A, "S", "K", True: v = GetSetting(A, "S", "K")                ' Boolean saved as text
    End Select
    On Error Resume Next: DeleteSetting A
    T = "SaveSetting." & k & "=" & Sh(v)
    Exit Function
fail:
    en = Err.Number
    Resume cleanup
cleanup:
    On Error Resume Next: DeleteSetting A
    T = "SaveSetting." & k & "=ERR" & en
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
    On Error Resume Next: DeleteSetting A
End Function
