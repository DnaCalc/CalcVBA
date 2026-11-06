Attribute VB_Name = "UdfRuntime"
Option Explicit
Private Declare PtrSafe Function FindWindowExA Lib "user32" (ByVal parent As LongPtr, ByVal after As LongPtr, ByVal className As String, ByVal title As LongPtr) As LongPtr
Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" (ByVal window As LongPtr, ByVal objectId As Long, ByRef iid As DispatchId, ByRef result As Object) As Long
Private Type DispatchId
    d0 As Long
    d1 As Long
    d2 As Long
    d3 As Long
End Type
Private cachedApplication As Object
'--- procedures ---
Private Function UdfCurrency(ByVal scaled As LongLong) As Currency
    CopyMemory VarPtr(UdfCurrency),VarPtr(scaled),8^
End Function
Public Property Get Application() As Object
    Dim hwnd As XLOPER12, pointers As OperPointers, code As Long
    Dim window As LongPtr, desk As LongPtr, grid As LongPtr
    Dim iid As DispatchId, native As Object
    If cachedApplication Is Nothing Then
        code=ExcelCall(16392&,0&,pointers,hwnd)
        If code<>0& Then Err.Raise 5&,"UdfRuntime.Application","xlGetHwnd failed"
        window=hwnd.payload0 And &HFFFFFFFF^
        ExcelFree hwnd
        desk=FindWindowExA(window,0^,"XLDESK",0^)
        If desk=0^ Then Err.Raise 5&,"UdfRuntime.Application","Excel desktop is unavailable"
        grid=FindWindowExA(desk,0^,"EXCEL7",0^)
        If grid=0^ Then Err.Raise 5&,"UdfRuntime.Application","Excel window is unavailable"
        iid.d0=&H20400: iid.d1=0&: iid.d2=&HC0: iid.d3=&H46000000
        code=AccessibleObjectFromWindow(grid,&HFFFFFFF0,iid,native)
        If code<>0& Then Err.Raise 5&,"UdfRuntime.Application","native object model is unavailable"
        Set cachedApplication=native.Application
    End If
    Set Application=cachedApplication
End Property

Private Function ReadOperand(ByVal pointer As LongPtr) As XLOPER12
    If pointer=0^ Then
        ReadOperand.xltype=128&
    Else
        CopyMemory VarPtr(ReadOperand),pointer,32^
    End If
End Function
Private Function UdfMissing(ByVal pointer As LongPtr) As Boolean
    Dim value As XLOPER12
    value=ReadOperand(pointer)
    UdfMissing=((value.xltype And 4095&)=128&)
End Function
Private Function UdfIsReference(ByVal pointer As LongPtr) As Boolean
    Dim value As XLOPER12, kind As Long
    value=ReadOperand(pointer)
    kind=value.xltype And 4095&
    UdfIsReference=(kind=8& Or kind=1024&)
End Function
Private Function UdfValue(ByVal pointer As LongPtr) As Variant
    Dim value As XLOPER12, converted As XLOPER12, pointers As OperPointers, code As Long
    value=ReadOperand(pointer)
    If UdfIsReference(pointer) Then
        pointers.p0=VarPtr(value)
        code=ExcelCall(16386&,1&,pointers,converted)
        If code<>0& Then Err.Raise 13&
        UdfValue=XlToVariant(converted)
        ExcelFree converted
    Else
        UdfValue=XlToVariant(value)
    End If
End Function
Private Function ColumnName(ByVal column As Long) As String
    column=column+1&
    Do While column>0&
        column=column-1&
        ColumnName=Chr$(65&+(column Mod 26&)) & ColumnName
        column=column\26&
    Loop
End Function
Private Function UdfRange(ByVal pointer As LongPtr) As Object
    Dim value As XLOPER12, fullReference As XLOPER12, requested As XLOPER12
    Dim name As XLOPER12, pointers As OperPointers, code As Long
    Dim sheet As String, address As String, rectangle As XlRectangle
    Dim index As Long, count As Long, area As Object, app As Object
    On Error GoTo Failed
    value=ReadOperand(pointer)
    If Not UdfIsReference(pointer) Then Err.Raise 13&
    If (value.xltype And 4095&)=1024& Then
        requested.payload0=8^: requested.xltype=2048&
        pointers.p0=VarPtr(value): pointers.p1=VarPtr(requested)
        code=ExcelCall(16386&,2&,pointers,fullReference)
        If code<>0& Then Err.Raise 13&
        value=fullReference
    End If
    pointers.p0=VarPtr(value)
    code=ExcelCall(16389&,1&,pointers,name)
    If code<>0& Then Err.Raise 13&
    sheet=CStr(XlToVariant(name))
    ExcelFree name
    count=XlReferenceCount(value)
    Set app=Application
    For index=0& To count-1&
        rectangle=XlReferenceArea(value,index)
        address="'" & Replace(sheet,"'","''") & "'!$" & ColumnName(rectangle.firstColumn) & "$" & CStr(rectangle.firstRow+1&) & ":$" & ColumnName(rectangle.lastColumn) & "$" & CStr(rectangle.lastRow+1&)
        Set area=app.Range(address)
        If UdfRange Is Nothing Then
            Set UdfRange=area
        Else
            Set UdfRange=app.Union(UdfRange,area)
        End If
    Next index
    ExcelFree fullReference
    Exit Function
Failed:
    ExcelFree name
    ExcelFree fullReference
    Err.Raise 13&
End Function
