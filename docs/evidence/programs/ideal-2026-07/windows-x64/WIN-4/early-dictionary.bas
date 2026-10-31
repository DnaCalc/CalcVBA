Option Explicit
Public Function RunProbe() As String
    On Error GoTo Failed
    Dim obj As Scripting.Dictionary, text As String, missing As Variant
    Set obj = New Scripting.Dictionary
    obj.Add "a", 19
    obj.Item("a") = 42
    text = "value=" & CStr(obj.Item("a")) & "|count=" & CStr(obj.Count)
    text = text & "|exists=" & CStr(obj.Exists("a"))
    missing = obj.Item("absent")
    text = text & "|missing=" & TypeName(missing) & "|after=" & CStr(obj.Count)
    On Error Resume Next
    obj.Add "a", 7
    text = text & "|duplicate=" & CStr(Err.Number) & "|source=" & Err.Source & "|description=" & Err.Description
    Err.Clear
    On Error GoTo Failed
    Set obj = Nothing
    RunProbe = text
    Exit Function
Failed:
    RunProbe = "FAILED:" & CStr(Err.Number) & ":" & Err.Description
End Function
