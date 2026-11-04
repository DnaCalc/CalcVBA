Attribute VB_Name = "PlainXll"
Option Explicit

' Ordinary VBA and Windows declarations. No compiler knowledge of Excel.
Private Declare PtrSafe Function GetModuleHandleW Lib "kernel32" (ByVal name As LongPtr) As LongPtr
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal moduleHandle As LongPtr, ByVal name As String) As LongPtr
Private Declare PtrSafe Function CoTaskMemAlloc Lib "ole32" (ByVal count As LongPtr) As LongPtr
Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal address As LongPtr)
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal destination As LongPtr, ByVal source As LongPtr, ByVal count As LongPtr)
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" (ByVal instance As LongPtr, ByVal address As LongPtr, ByVal convention As Long, ByVal returnType As Integer, ByVal count As Long, ByVal types As LongPtr, ByVal arguments As LongPtr, ByRef result As Variant) As Long

' x64 XLOPER12: 24-byte union, DWORD type, four bytes trailing padding.
Public Type XLOPER12
    payload0 As LongLong
    payload1 As LongLong
    payload2 As LongLong
    xltype As Long
    padding As Long
End Type

Private Type OperPointers
    p0 As LongPtr
    p1 As LongPtr
    p2 As LongPtr
    p3 As LongPtr
    p4 As LongPtr
    p5 As LongPtr
    p6 As LongPtr
End Type

Private Type CallTypes
    t0 As Integer
    t1 As Integer
    t2 As Integer
    t3 As Integer
End Type

Private Type CallValues
    p0 As LongPtr
    p1 As LongPtr
    p2 As LongPtr
    p3 As LongPtr
End Type

Private Type RegistrationRow
    exportName As String
    typeText As String
    worksheetName As String
    argumentNames As String
End Type

Private callback As LongPtr
Private outstanding As Long
Private firstRegistration As XLOPER12
Private secondRegistration As XLOPER12

Private Function ExcelCall(ByVal functionNumber As Long, ByVal count As Long, ByRef args As OperPointers, ByRef result As XLOPER12) As Long
    Dim types As CallTypes, values As CallValues
    Dim argsPointer As LongPtr, resultPointer As LongPtr
    Dim status As Variant, hr As Long
    Dim v0 As Variant, v1 As Variant, v2 As Variant, v3 As Variant
    argsPointer = VarPtr(args)
    resultPointer = VarPtr(result)
    types.t0 = 3
    types.t1 = 3
    types.t2 = 20
    types.t3 = 20
    ' DispCallFunc takes pointers to VARIANTARGs, including their headers.
    v0 = functionNumber
    v1 = count
    v2 = argsPointer
    v3 = resultPointer
    values.p0 = VarPtr(v0)
    values.p1 = VarPtr(v1)
    values.p2 = VarPtr(v2)
    values.p3 = VarPtr(v3)
    hr = DispCallFunc(0^, callback, 4&, 3, 4&, VarPtr(types), VarPtr(values), status)
    If hr <> 0& Then Err.Raise 5&, "PlainXll.ExcelCall", "DispCallFunc failed"
    ExcelCall = CLng(status)
End Function

Private Function NumberOper(ByVal value As Double) As XLOPER12
    CopyMemory VarPtr(NumberOper), VarPtr(value), 8^
    NumberOper.xltype = 1&
End Function

Private Function StringOper(ByVal value As String) As XLOPER12
    Dim length As Integer, data As LongPtr
    length = CInt(Len(value))
    data = CoTaskMemAlloc(CLngPtr(length) * 2^ + 4^)
    If data = 0^ Then Err.Raise 7&
    CopyMemory data, VarPtr(length), 2^
    CopyMemory data + 2^, StrPtr(value), CLngPtr(length) * 2^
    StringOper.payload0 = data
    StringOper.xltype = 2&
End Function

Private Sub FreeStringOper(ByRef value As XLOPER12)
    CoTaskMemFree value.payload0
    value.payload0 = 0^
End Sub

' Registration table kept in VBA: export, C API types, worksheet name, arguments.
Private Function RegistrationTable(ByVal index As Long) As RegistrationRow
    If index = 1& Then
        RegistrationTable.exportName = "AddDouble"
        RegistrationTable.typeText = "BBB"
        RegistrationTable.worksheetName = "CalcVbaAdd"
        RegistrationTable.argumentNames = "left,right"
    ElseIf index = 2& Then
        RegistrationTable.exportName = "NativeTwice"
        RegistrationTable.typeText = "QQ"
        RegistrationTable.worksheetName = "CalcVbaTwice"
        RegistrationTable.argumentNames = "value"
    Else
        Err.Raise 9&
    End If
End Function

Private Function RegisterRow(ByRef dllName As XLOPER12, ByVal index As Long) As XLOPER12
    Dim row As RegistrationRow
    row = RegistrationTable(index)
    RegisterRow = RegisterFunction(dllName, CStr(row.exportName), CStr(row.typeText), CStr(row.worksheetName), CStr(row.argumentNames))
End Function

Private Function RegisterFunction(ByRef dllName As XLOPER12, ByVal exportName As String, ByVal types As String, ByVal worksheetName As String, ByVal argumentNames As String) As XLOPER12
    Dim proc As XLOPER12, typeText As XLOPER12, name As XLOPER12
    Dim arguments As XLOPER12, macroType As XLOPER12, category As XLOPER12
    Dim pointers As OperPointers, code As Long
    proc = StringOper(exportName)
    typeText = StringOper(types)
    name = StringOper(worksheetName)
    arguments = StringOper(argumentNames)
    category = StringOper("CalcVBA")
    macroType = NumberOper(1#)
    pointers.p0 = VarPtr(dllName)
    pointers.p1 = VarPtr(proc)
    pointers.p2 = VarPtr(typeText)
    pointers.p3 = VarPtr(name)
    pointers.p4 = VarPtr(arguments)
    pointers.p5 = VarPtr(macroType)
    pointers.p6 = VarPtr(category)
    code = ExcelCall(149&, 7&, pointers, RegisterFunction)
    FreeStringOper proc
    FreeStringOper typeText
    FreeStringOper name
    FreeStringOper arguments
    FreeStringOper category
    If code <> 0& Then Err.Raise 5&, "PlainXll.RegisterFunction", "xlfRegister failed"
End Function

Public Function xlAutoOpen() As Long
    Dim dllName As XLOPER12, unused As XLOPER12, pointers As OperPointers
    Dim code As Long
    callback = GetProcAddress(GetModuleHandleW(0^), "MdCallBack12")
    If callback = 0^ Then Err.Raise 513&, "PlainXll.xlAutoOpen", "Excel callback missing"
    code = ExcelCall(16393&, 0&, pointers, dllName)
    If code <> 0& Then Err.Raise 514& + code, "PlainXll.xlAutoOpen", "xlGetName failed"
    firstRegistration = RegisterRow(dllName, 1&)
    secondRegistration = RegisterRow(dllName, 2&)
    pointers.p0 = VarPtr(dllName)
    code = ExcelCall(16384&, 1&, pointers, unused)
    If firstRegistration.xltype <> 1& Then Err.Raise 600& + firstRegistration.xltype, "PlainXll.xlAutoOpen", "first registration failed"
    If secondRegistration.xltype <> 1& Then Err.Raise 700& + secondRegistration.xltype, "PlainXll.xlAutoOpen", "second registration failed"
    xlAutoOpen = 1&
End Function

Public Function xlAutoClose() As Long
    Dim pointers As OperPointers, unused As XLOPER12, code As Long
    pointers.p0 = VarPtr(firstRegistration)
    code = ExcelCall(201&, 1&, pointers, unused)
    pointers.p0 = VarPtr(secondRegistration)
    code = ExcelCall(201&, 1&, pointers, unused)
    xlAutoClose = 1&
End Function

Public Function AddDouble(ByVal left As Double, ByVal right As Double) As Double
    AddDouble = left + right
End Function

Public Function NativeTwice(ByVal operand As LongPtr) As LongPtr
    Dim inputValue As XLOPER12, result As XLOPER12, number As Double
    CopyMemory VarPtr(inputValue), operand, 32^
    If (inputValue.xltype And 4095&) = 1& Then
        CopyMemory VarPtr(number), VarPtr(inputValue), 8^
        result = NumberOper(number * 2#)
    Else
        result.payload0 = 15^
        result.xltype = 16&
    End If
    result.xltype = result.xltype Or 16384&
    NativeTwice = CoTaskMemAlloc(32^)
    If NativeTwice = 0^ Then Err.Raise 7&
    CopyMemory NativeTwice, VarPtr(result), 32^
    outstanding = outstanding + 1&
End Function

Public Sub xlAutoFree12(ByVal operand As LongPtr)
    CoTaskMemFree operand
    outstanding = outstanding - 1&
End Sub

Public Function OutstandingResults() As Long
    OutstandingResults = outstanding
End Function

' Runs through the same pointer/Declare operations on VM3, JIT and AOT.
Public Function ConversionProbe(ByVal value As Double) As Double
    Dim record As XLOPER12
    record = NumberOper(value)
    CopyMemory VarPtr(ConversionProbe), VarPtr(record), 8^
End Function

Public Sub Main()
End Sub
