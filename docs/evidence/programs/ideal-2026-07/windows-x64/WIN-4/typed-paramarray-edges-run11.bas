Option Explicit
Private Sub Touch(ParamArray items() As Variant)
    items(0) = 9
End Sub
Private Function AliasRead(ParamArray items() As Variant) As Long
    items(0) = 9
    AliasRead = items(1)
End Function
Private Function CopyTail(ParamArray items() As Variant) As String
    Dim copied As Variant
    copied = items
    copied(0) = 27
    CopyTail = items(0) & ":" & copied(0)
End Function
Private Sub ClearObject(ParamArray items() As Variant)
    Set items(0) = Nothing
End Sub
Private Function BadWrite(ParamArray items() As Variant) As String
    On Error Resume Next
    items(0) = "bad"
    BadWrite = Err.Number & ":" & items(0)
End Function
Private Function Nesting(ParamArray items() As Variant) As String
    Nesting = LBound(items) & ":" & UBound(items) & ":" & IsArray(items(0))
End Function
Public Function Observe(ByVal obj As Object) As String
    Dim x As Long, result As Long, a As Variant
    x = 5
    Touch (x)
    Observe = "private=" & x
    result = AliasRead(x, x)
    Observe = Observe & "|repeat=" & result & ":" & x
    Observe = Observe & "|copy=" & CopyTail(x) & ":" & x
    Observe = Observe & "|bad=" & BadWrite(x) & ":" & x
    a = Array(2, 3)
    Observe = Observe & "|nested=" & Nesting(a)
    ClearObject obj
    Observe = Observe & "|object=" & (obj Is Nothing)
End Function
Public Function RunProbe() As String
    Dim obj As Object
    Set obj = CreateObject("Scripting.Dictionary")
    RunProbe = Observe(obj)
End Function
