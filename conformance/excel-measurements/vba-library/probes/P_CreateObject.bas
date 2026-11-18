Attribute VB_Name = "P_CreateObject"
Option Explicit
' CreateObject (Interaction; only Scripting.Dictionary and error cases).
Private Const N As Long = 19

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    Dim o As Object
    On Error GoTo fail
    Select Case k
        Case 1: Set o = CreateObject("Scripting.Dictionary"): v = TypeName(o)             ' TypeName of the object
        Case 2: Set o = CreateObject("Scripting.Dictionary"): v = o.Count                 ' new Dictionary is empty
        Case 3: Set o = CreateObject("Scripting.Dictionary"): v = IsObject(o)             ' an object
        Case 4: Set o = CreateObject("Scripting.Dictionary"): v = VarType(o)              ' vbObject
        Case 5: Set o = CreateObject("Scripting.Dictionary"): o.Add "a", 1: v = o.Exists("a") & "," & o("a")   ' usable
        Case 6: Set o = CreateObject("Scripting.Dictionary"): v = o.CompareMode           ' default binary compare
        Case 7: Set o = CreateObject("scripting.dictionary"): v = TypeName(o)             ' ProgID case-insensitive
        Case 8: Set o = CreateObject("Scripting.Dictionary", ""): v = TypeName(o)         ' empty server name
        Case 9: Set o = CreateObject(Class:="Scripting.Dictionary"): v = TypeName(o)      ' named argument
        Case 10: Set o = CreateObject("Scripting.Dictionary.1"): v = TypeName(o)           ' version-dependent ProgID
        Case 11: Set o = CreateObject("CalcVbaProbe.NoSuchClass"): v = TypeName(o)         ' bad ProgID: 429
        Case 12: Set o = CreateObject(""): v = TypeName(o)                                 ' empty ProgID
        Case 13: Set o = CreateObject(" Scripting.Dictionary"): v = TypeName(o)            ' leading space
        Case 14: Set o = CreateObject("Scripting.Dictionary "): v = TypeName(o)            ' trailing space
        Case 15: Set o = CreateObject("Scripting"): v = TypeName(o)                        ' library name only
        Case 16: Set o = CreateObject(Null): v = TypeName(o)                               ' Null class: error 94
        Case 17: Set o = CreateObject("NoDots"): v = TypeName(o)                           ' ProgID without a dot
        Case 18: Set o = CreateObject("Scripting.Dictionary"): v = o Is Nothing            ' not Nothing
        Case 19: Set o = CreateObject("Scripting.Dictionary"): v = TypeName(CreateObject("Scripting.Dictionary")) & "|" & (o Is CreateObject("Scripting.Dictionary"))   ' each call a new object
    End Select
    T = "CreateObject." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "CreateObject." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
