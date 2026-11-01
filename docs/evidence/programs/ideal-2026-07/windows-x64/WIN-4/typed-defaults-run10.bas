Option Explicit

Private Function AddValues(ByVal x As Variant, ByVal y As Variant) As Variant
    AddValues = x + y
End Function
Private Function TextValue(ByVal x As Variant) As String
    TextValue = "v=" & x
End Function
Private Function NumberValue(ByVal x As Variant) As Double
    NumberValue = x
End Function
Private Function EqualValue(ByVal x As Variant) As Boolean
    EqualValue = x = 41#
End Function
Private Function ConditionValue(ByVal x As Variant) As Long
    If x Then ConditionValue = 1
End Function
Private Function ErrorValue(ByVal x As Variant) As Long
    On Error GoTo Failed
    Dim n As Double
    n = x
    Exit Function
Failed:
    ErrorValue = Err.Number
End Function
Public Function RunProbe() As String
    On Error GoTo Failed
    Dim x As Variant, y As Variant, absent As Object, c As Collection
    Set x = ThisWorkbook.Worksheets(1).Range("A1")
    Set y = ThisWorkbook.Worksheets(1).Range("A2")
    x.Value2 = 41#
    y.Value2 = 1#
    Dim s As String
    s = "add=" & CStr(AddValues(x, y))
    s = s & "|text=" & TextValue(x)
    s = s & "|number=" & CStr(NumberValue(x))
    s = s & "|equal=" & CStr(EqualValue(x))
    s = s & "|condition=" & CStr(ConditionValue(x))
    s = s & "|identity=" & CStr(IsObject(x))
    s = s & "|nothing=" & CStr(ErrorValue(absent))
    s = s & "|null=" & CStr(ErrorValue(Null))
    x.Value2 = "bad"
    s = s & "|bad=" & CStr(ErrorValue(x))
    Set c = New Collection
    s = s & "|required-default=" & CStr(ErrorValue(c))
    RunProbe = s
    Exit Function
Failed:
    RunProbe = "FAILED:" & CStr(Err.Number) & ":" & Err.Description
End Function
