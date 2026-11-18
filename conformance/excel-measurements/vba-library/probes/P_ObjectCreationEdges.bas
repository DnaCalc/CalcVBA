Attribute VB_Name = "P_ObjectCreationEdges"
Option Explicit
' UNMEASURED local/remote activation and shadowing edges.
Private Const N As Long = 19
Private Function CreateObject(ByVal Class As Variant, Optional ByVal ServerName As Variant) As Long
    CreateObject = 42
End Function
Private Function GetObject(Optional ByVal PathName As Variant, Optional ByVal Class As Variant) As Long
    GetObject = 43
End Function
Private Function T(ByVal k As Long) As String
    Dim v As Variant, o As Object, other As Object
    Dim class As String, server As String, path As String
    On Error GoTo fail
    class = "Scripting." & "Dictionary"
    Select Case k
        Case 1: Set o = VBA.CreateObject(class): v = TypeName(o) & ":" & o.Count
        Case 2: Set o = VBA.CreateObject(class, server): v = TypeName(o) & ":" & o.Count
        Case 3: Set o = VBA.CreateObject(Class:=class): v = o.Count
        Case 4: Set o = VBA.GetObject(path, class): v = TypeName(o) & ":" & o.Count
        Case 5: Set o = VBA.CreateObject(class): Set other = VBA.CreateObject(class): v = o Is other
        Case 6: Set o = VBA.CreateObject(class): v = VarType(o)
        Case 7: Set o = VBA.CreateObject(123): v = TypeName(o)
        Case 8: Set o = VBA.CreateObject(Empty): v = TypeName(o)
        Case 9: Set o = VBA.CreateObject(Null): v = TypeName(o)
        Case 10: Set o = VBA.CreateObject(class, Null): v = o.Count
        Case 11: Set o = VBA.CreateObject(class, Empty): v = o.Count
        Case 12: Set o = VBA.CreateObject(class, "."): v = o.Count
        Case 13: Set o = VBA.CreateObject(class, "localhost"): v = o.Count
        Case 14: Set o = VBA.CreateObject(class, 0): v = o.Count
        Case 15: Set o = VBA.GetObject(Empty, class): v = o.Count
        Case 16: Set o = VBA.GetObject("", Empty): v = TypeName(o)
        Case 17: Set o = VBA.GetObject("", Null): v = TypeName(o)
        Case 18: v = CreateObject("Scripting.Dictionary")
        Case 19: v = GetObject("", "Scripting.Dictionary")
    End Select
    T = "ObjectCreationEdges." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "ObjectCreationEdges." & k & "=ERR" & Err.Number
End Function
Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
