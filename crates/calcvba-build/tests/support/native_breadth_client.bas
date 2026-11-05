Option Explicit
Private Declare PtrSafe Function NStringSuite Lib "__DLL__" Alias "StringSuite" () As LongPtr
Private Declare PtrSafe Function NVariantSuite Lib "__DLL__" Alias "VariantSuite" () As LongPtr
Private Declare PtrSafe Function NLibrarySuite Lib "__DLL__" Alias "LibrarySuite" () As LongPtr
Private Declare PtrSafe Function NArraySuite Lib "__DLL__" Alias "ArraySuite" () As LongPtr
Private Declare PtrSafe Function NErrorSuite Lib "__DLL__" Alias "ErrorSuite" () As LongPtr
Private Declare PtrSafe Function NResumeSuite Lib "__DLL__" Alias "ResumeSuite" () As LongPtr
Private Declare PtrSafe Function NNestedErrorSuite Lib "__DLL__" Alias "NestedErrorSuite" () As LongPtr
Private Declare PtrSafe Function NMixedSuite Lib "__DLL__" Alias "MixedSuite" () As LongPtr
Private Declare PtrSafe Function NArrayAllocationSuite Lib "__DLL__" Alias "ArrayAllocationSuite" () As LongPtr
Private Declare PtrSafe Function NEraseSuite Lib "__DLL__" Alias "EraseSuite" () As LongPtr
Private Declare PtrSafe Function NErrorBoundsSuite Lib "__DLL__" Alias "ErrorBoundsSuite" () As LongPtr
Private Declare PtrSafe Function NActivationSuite Lib "__DLL__" Alias "ActivationSuite" () As LongPtr
Private Declare PtrSafe Function NForeignObjectSuite Lib "__DLL__" Alias "ForeignObjectSuite" (ByVal foreign As LongPtr) As LongPtr
Private Declare PtrSafe Function NShutdown Lib "__DLL__" Alias "CalcVbaShutdown" () As Long
Private Declare PtrSafe Function NLastError Lib "__DLL__" Alias "CalcVbaGetLastError" (ByRef source As LongPtr, ByRef description As LongPtr) As Long
Private Declare PtrSafe Function BStrLength Lib "oleaut32" Alias "SysStringLen" (ByVal value As LongPtr) As Long
Private Declare PtrSafe Sub FreeBStr Lib "oleaut32" Alias "SysFreeString" (ByVal value As LongPtr)
Private Declare PtrSafe Sub CopyBytes Lib "kernel32" Alias "RtlMoveMemory" (ByVal destination As LongPtr, ByVal source As LongPtr, ByVal count As LongPtr)
Private Function TakeNativeString(ByVal pointer As LongPtr) As String
    Dim length As Long
    length=BStrLength(pointer)
    TakeNativeString=String$(length, ChrW$(0))
    If length Then CopyBytes StrPtr(TakeNativeString), pointer, length*2&
    FreeBStr pointer
End Function
Private Sub CheckNative(ByVal name As String, ByVal actual As String, ByVal expected As String)
    Dim number As Long, source As LongPtr, description As LongPtr, detail As String
    number=NLastError(source, description)
    detail=TakeNativeString(source) & ":" & TakeNativeString(description)
    If number<>0 Then Err.Raise 543, name, "native error " & CStr(number) & " " & detail
    If actual<>expected Then Err.Raise 544, name, "actual=" & actual & "; expected=" & expected
End Sub
Public Function ProbeNativeExports() As String
    Dim dictionary As Object, cell As Object, result As String
    Set dictionary=CreateObject("Scripting.Dictionary")
    dictionary.Add "owned-client", 42&
    CheckNative "StringSuite", TakeNativeString(NStringSuite()), StringSuite()
    CheckNative "VariantSuite", TakeNativeString(NVariantSuite()), VariantSuite()
    CheckNative "LibrarySuite", TakeNativeString(NLibrarySuite()), LibrarySuite()
    CheckNative "ArraySuite", TakeNativeString(NArraySuite()), ArraySuite()
    CheckNative "ErrorSuite", TakeNativeString(NErrorSuite()), ErrorSuite()
    CheckNative "ResumeSuite", TakeNativeString(NResumeSuite()), ResumeSuite()
    CheckNative "NestedErrorSuite", TakeNativeString(NNestedErrorSuite()), NestedErrorSuite()
    CheckNative "MixedSuite", TakeNativeString(NMixedSuite()), MixedSuite()
    CheckNative "ArrayAllocationSuite", TakeNativeString(NArrayAllocationSuite()), ArrayAllocationSuite()
    CheckNative "EraseSuite", TakeNativeString(NEraseSuite()), EraseSuite()
    CheckNative "ErrorBoundsSuite", TakeNativeString(NErrorBoundsSuite()), ErrorBoundsSuite()
    CheckNative "ActivationSuite", TakeNativeString(NActivationSuite()), ActivationSuite()
    CheckNative "ForeignDictionary", TakeNativeString(NForeignObjectSuite(ObjPtr(dictionary))), ForeignObjectSuite(dictionary)
    CheckNative "OwnedExcel", TakeNativeString(NForeignObjectSuite(ObjPtr(Application))), ForeignObjectSuite(Application)
    Set cell=ThisWorkbook.Worksheets(1).Range("A1")
    cell.Value2=37#
    CheckNative "OwnedExcelRange", TakeNativeString(NForeignObjectSuite(ObjPtr(cell))), ForeignObjectSuite(cell)
    Set cell=ThisWorkbook.Worksheets(1).Range("A1:A2")
    cell.Cells(2, 1).Value2="array item"
    CheckNative "OwnedExcelArray", TakeNativeString(NForeignObjectSuite(ObjPtr(cell))), ForeignObjectSuite(cell)
    Set cell=Nothing
    If dictionary.Item("owned-client")<>42& Then Err.Raise 545
    If Application.Name<>"Microsoft Excel" Then Err.Raise 546
    Set dictionary=Nothing
    If NShutdown()<>0 Then Err.Raise 547, , "native shutdown failed"
    ProbeNativeExports="13 native exports|Dictionary activation|foreign Dictionary|owned Excel and Range|shutdown"
End Function
