Attribute VB_Name = "OracleSelfTest"
Option Explicit
Private Declare PtrSafe Function CloseXll Lib "@XLL_PATH@" Alias "xlAutoClose" () As Long
Public Function ProbeNativeExports() As String
    Dim sheet As Worksheet, index As Long, beforeVba As Long, beforeNative As Long
    Dim ordinaryVba As Long, ordinaryNative As Long, code As Long
    On Error GoTo Failed
    If Not Application.RegisterXLL("@XLL_PATH@") Then Err.Raise 513&,"UDF oracle","RegisterXLL failed"
    Set sheet=ActiveWorkbook.Worksheets(1)
    sheet.Name="Data's table"
    sheet.Range("A1").Value2=2
    sheet.Range("A2").Value2=3
    Compare sheet,3,"Scalar(""2"",3)"
    Compare sheet,4,"RangeSum(A1:A2)"
    Compare sheet,5,"ReferenceKind(A1:A2)"
    Compare sheet,6,"Defaults()"
    Compare sheet,7,"Defaults(,9)"
    Compare sheet,8,"SumItems(1,2,3,4)"
    Compare sheet,9,"ErrorValue()"
    Compare sheet,10,"Unhandled()"
    Compare sheet,11,"CallerRow()"
    Compare sheet,12,"SumItems(1)"
    Compare sheet,13,"ReferenceKind(5)"
    Compare sheet,14,"ArrayArgument({1,2;3,4})"
    Compare sheet,15,"ReferenceKind((A1:A2,A4:A5))"
    Compare sheet,16,"OptionalReference()"
    Compare sheet,17,"OptionalReference(A1:A2)"
    Compare sheet,18,"OptionalData()"
    Compare sheet,19,"ScalarKinds(""True"",3.5,9)"
    sheet.Range("B20:C21").FormulaArray="=Matrix({1,2;3,4})"
    sheet.Range("E20:F21").FormulaArray="=Native_Matrix({1,2;3,4})"
    For index=0 To 3
        AssertSame sheet.Cells(20+index\2,2+index Mod 2).Value2,sheet.Cells(20+index\2,5+index Mod 2).Value2,"Matrix"
    Next index
    sheet.Range("B30").Formula="=VolatileCount()"
    sheet.Range("E30").Formula="=Native_VolatileCount()"
    sheet.Range("B31").Formula="=OrdinaryCount()"
    sheet.Range("E31").Formula="=Native_OrdinaryCount()"
    Application.Calculate
    beforeVba=CLng(sheet.Range("B30").Value2)
    beforeNative=CLng(sheet.Range("E30").Value2)
    ordinaryVba=CLng(sheet.Range("B31").Value2)
    ordinaryNative=CLng(sheet.Range("E31").Value2)
    Application.Calculate
    If CLng(sheet.Range("B30").Value2)-beforeVba<>1 Or CLng(sheet.Range("E30").Value2)-beforeNative<>1 Then Err.Raise 513&,"UDF oracle","Volatile did not recalculate once"
    If CLng(sheet.Range("B31").Value2)<>ordinaryVba Or CLng(sheet.Range("E31").Value2)<>ordinaryNative Then Err.Raise 513&,"UDF oracle","Ordinary function became volatile"
    code=CloseXll()
    ProbeNativeExports="scalar,Range,Variant reference,Optional,ParamArray,array,CVErr,error,Caller,Volatile: matched"
    Exit Function
Failed:
    Dim message As String
    message=Err.Source & ":" & Err.Description
    On Error Resume Next
    code=CloseXll()
    On Error GoTo 0
    Err.Raise 513&,"UDF oracle",message
End Function
Private Sub Compare(ByVal sheet As Worksheet,ByVal row As Long,ByVal formula As String)
    sheet.Cells(row,2).Formula="=" & formula
    sheet.Cells(row,5).Formula="=Native_" & formula
    sheet.Calculate
    AssertSame sheet.Cells(row,2).Value2,sheet.Cells(row,5).Value2,formula
End Sub
Private Sub AssertSame(ByVal ordinary As Variant,ByVal native As Variant,ByVal label As String)
    If IsError(ordinary) Then
        If Not IsError(native) Then Err.Raise 513&,"UDF oracle",label & " expected error, got " & CStr(native)
        If CStr(ordinary)<>CStr(native) Then Err.Raise 513&,"UDF oracle",label & " error mismatch: " & CStr(ordinary) & ":" & CStr(native)
    ElseIf IsError(native) Then
        Err.Raise 513&,"UDF oracle",label & " native error: " & CStr(native)
    ElseIf CStr(ordinary)<>CStr(native) Then
        Err.Raise 513&,"UDF oracle",label & " mismatch: " & CStr(ordinary) & ":" & CStr(native)
    End If
End Sub
