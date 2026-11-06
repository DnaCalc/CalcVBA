Attribute VB_Name = "Functions"
Option Explicit
Private volatileCalls As Long
Private ordinaryCalls As Long
Public Function Scalar(ByVal value As Double, ByVal increment As Long) As Double
    Dim dictionary As New Scripting.Dictionary
    dictionary.Add "value",value
    Scalar=CDbl(dictionary.Item("value"))+increment
End Function
Public Function RangeSum(ByVal cells As Excel.Range) As Double
    RangeSum=CDbl(cells.Cells(1,1).Value2)+CDbl(cells.Cells(2,1).Value2)
End Function
Public Function ReferenceKind(ByVal value As Variant) As String
    If IsObject(value) Then
        ReferenceKind=TypeName(value) & ":" & CStr(value.Rows.Count)
    Else
        ReferenceKind=TypeName(value)
    End If
End Function
Public Function Defaults(Optional ByVal number As Long=7, Optional value As Variant) As String
    Defaults=CStr(number) & ":" & CStr(IsMissing(value))
End Function
Public Function SumItems(ByVal first As Double, ParamArray items() As Variant) As Double
    Dim index As Long
    SumItems=first
    For index=LBound(items) To UBound(items)
        SumItems=SumItems+CDbl(items(index))
    Next index
End Function
Public Function Matrix(ByVal value As Variant) As Variant
    Dim result(1 To 2,1 To 2) As Variant
    result(1,1)=value(1,1)
    result(1,2)="matrix"
    result(2,1)=True
    result(2,2)=CVErr(2042)
    Matrix=result
End Function
Public Function ErrorValue() As Variant
    ErrorValue=CVErr(2042)
End Function
Public Function Unhandled() As Variant
    Err.Raise 5&
End Function
Public Function CallerRow() As Long
    CallerRow=Application.Caller.Row
End Function
Public Function VolatileCount() As Long
    Application.Volatile True
    volatileCalls=volatileCalls+1&
    VolatileCount=volatileCalls
End Function
Public Function OrdinaryCount() As Long
    ordinaryCalls=ordinaryCalls+1&
    OrdinaryCount=ordinaryCalls
End Function
Public Function ArrayArgument(ByRef items() As Variant) As Double
    ArrayArgument=CDbl(items(1,1))+CDbl(items(2,2))
End Function
Public Function OptionalReference(Optional ByVal cells As Excel.Range=Nothing) As String
    If cells Is Nothing Then
        OptionalReference="none"
    Else
        OptionalReference=TypeName(cells)
    End If
End Function
Public Function OptionalData(Optional ByVal text As String=vbCrLf,Optional ByVal amount As Currency=123.4567@) As String
    OptionalData=CStr(Len(text)) & ":" & CStr(amount)
End Function
Public Function ScalarKinds(ByVal flag As Boolean,ByVal number As Integer,ByVal text As String) As String
    ScalarKinds=CStr(flag) & ":" & CStr(number) & ":" & text
End Function
