Sub Main()
Dim target
Call Capture(target, 3, 4)
End Sub

Sub Capture(ByRef target, ParamArray items() As Variant)
target = UBound(items)
End Sub
