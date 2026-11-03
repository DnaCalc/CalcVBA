Public Function Version(ByVal app As Excel.Application) As String
    Version = app.Version
End Function
Public Function Build(ByVal app As Excel.Application) As Long
    Build = app.Build
End Function
Public Function Height(ByVal app As Excel.Application) As Double
    Height = app.Height
End Function
Public Function Visible(ByVal app As Excel.Application) As Boolean
    Visible = app.Visible
End Function
Public Function ReadRange(ByVal value As Excel.Range) As Variant
    ReadRange = value.value
End Function
Public Sub WriteRange(ByVal target As Excel.Range, ByVal value As Variant)
    target.value = value
End Sub
Public Function ReadValue2(ByVal value As Excel.Range) As Variant
    ReadValue2 = value.Value2
End Function
Public Sub WriteValue2(ByVal target As Excel.Range, ByVal value As Variant)
    target.Value2 = value
End Sub
Public Function ReadDefault(ByVal value As Excel.Range) As Variant
    ReadDefault = value
End Function
Public Sub WriteDefault(ByVal target As Excel.Range, ByVal value As Variant)
    target = value
End Sub
Public Function RuntimeDefault(ByVal value As Object) As Double
    RuntimeDefault = value + 1#
End Function
Public Function GetFunctions(ByVal app As Excel.Application) As Excel.WorksheetFunction
    Set GetFunctions = app.WorksheetFunction
End Function
Public Function SumArray(ByVal functions As Excel.WorksheetFunction, ByVal value As Variant) As Double
    SumArray = functions.Sum(value)
End Function
Public Function SumDifferentArities(ByVal functions As Excel.WorksheetFunction, ByVal x As Variant, ByVal y As Variant) As Double
    SumDifferentArities = functions.Sum(x) + functions.Sum(x, y)
End Function
Public Function SumRange(ByVal app As Excel.Application, ByVal value As Excel.Range) As Double
    SumRange = app.WorksheetFunction.Sum(value)
End Function
Public Function AverageArray(ByVal functions As Excel.WorksheetFunction, ByVal value As Variant) As Double
    AverageArray = functions.Average(value)
End Function
Public Function SumTyped(ByVal functions As Excel.WorksheetFunction, ByVal first As Long, ByVal second As Double) As Double
    SumTyped = functions.Sum(first, second)
End Function
Public Function MissingMatch(ByVal functions As Excel.WorksheetFunction, ByVal packet As Variant) As Long
    On Error GoTo Failed
    Dim result As Double
    result = functions.Match("missing", packet, 0#)
    MissingMatch = 0
    Exit Function
Failed:
    MissingMatch = Err.Number
End Function
Public Function ScalarKinds(ByVal target As Excel.Range) As String
    target.NumberFormat = "yyyy-mm-dd"
    target.Value2 = 45000#
    ScalarKinds = VarType(target.value) & ":" & VarType(target.Value2)
    target.NumberFormat = "$0.00"
    target.Value2 = 1.234567
    ScalarKinds = ScalarKinds & "|" & VarType(target.value) & ":" & VarType(target.Value2)
    target.NumberFormat = "General"
End Function
