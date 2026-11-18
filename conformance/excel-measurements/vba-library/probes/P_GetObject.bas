Attribute VB_Name = "P_GetObject"
Option Explicit
' GetObject (Interaction; only no-argument and error cases, plus Scripting.Dictionary with an empty path).
Private Const N As Long = 15

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    Dim o As Object
    On Error GoTo fail
    Select Case k
        Case 1: Set o = GetObject(): v = TypeName(o)                                      ' no arguments: error
        Case 2: Set o = GetObject(, "Scripting.Dictionary"): v = TypeName(o)              ' no running instance: 429
        Case 3: Set o = GetObject("", "Scripting.Dictionary"): v = TypeName(o)            ' empty path creates a new instance
        Case 4: Set o = GetObject(vbNullString, "Scripting.Dictionary"): v = TypeName(o)  ' vbNullString path
        Case 5: Set o = GetObject(Class:="CalcVbaProbe.NoSuchClass"): v = TypeName(o)     ' bad class, named argument
        Case 6: Set o = GetObject(, "CalcVbaProbe.NoSuchClass"): v = TypeName(o)          ' bad class: 429
        Case 7: Set o = GetObject("", "CalcVbaProbe.NoSuchClass"): v = TypeName(o)        ' empty path and bad class
        Case 8: Set o = GetObject(Environ("TEMP") & "\cvp-lib\no-such-file.cvp"): v = TypeName(o)   ' missing file: 432
        Case 9: Set o = GetObject(Null): v = TypeName(o)                                  ' Null path
        Case 10: Set o = GetObject(Empty): v = TypeName(o)                                 ' Empty path
        Case 11: Set o = GetObject(, ""): v = TypeName(o)                                  ' empty class
        Case 12: Set o = GetObject("", "Scripting.Dictionary"): v = o.Count                ' usable object
        Case 13: Set o = GetObject("", "Scripting.Dictionary"): v = IsObject(o)
        Case 14: Set o = GetObject(, Null): v = TypeName(o)                                ' Null class
        Case 15: Set o = GetObject(PathName:="", Class:="Scripting.Dictionary"): v = TypeName(o)   ' named arguments
    End Select
    T = "GetObject." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "GetObject." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
