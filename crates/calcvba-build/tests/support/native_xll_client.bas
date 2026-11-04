Option Explicit
Private Declare PtrSafe Function CloseXll Lib "__DLL__" Alias "xlAutoClose" () As Long
Private Declare PtrSafe Function AddDouble Lib "__DLL__" (ByVal left As Double, ByVal right As Double) As Double
Private Declare PtrSafe Function ConversionProbe Lib "__DLL__" (ByVal value As Double) As Double
Private Declare PtrSafe Function NativeTwice Lib "__DLL__" (ByVal operand As LongPtr) As LongPtr
Private Declare PtrSafe Sub xlAutoFree12 Lib "__DLL__" (ByVal operand As LongPtr)
Private Declare PtrSafe Function OutstandingResults Lib "__DLL__" () As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal destination As LongPtr, ByVal source As LongPtr, ByVal count As LongPtr)
Private Type XLOPER12
    number As Double
    unused1 As LongLong
    unused2 As LongLong
    xltype As Long
    padding As Long
End Type
Public Function ProbeNativeExports() As String
    If CloseXll() <> 1& Then Err.Raise 5&, "XllProbe", "empty registration table close failed"
    Dim inputValue As XLOPER12, result As XLOPER12, address As LongPtr
    Dim observed As String
    observed = CStr(AddDouble(2.25#, 3.5#)) & "|" & CStr(ConversionProbe(3.75#))
    inputValue.number = 4.5#
    inputValue.xltype = 1&
    address = NativeTwice(VarPtr(inputValue))
    CopyMemory VarPtr(result), address, 32^
    observed = observed & "|" & CStr(result.number) & "|" & CStr(result.xltype) & "|" & CStr(OutstandingResults())
    xlAutoFree12 address
    ProbeNativeExports = observed & "|" & CStr(OutstandingResults())
End Function
Public Sub Main()
    Dim observed As String
    observed = ProbeNativeExports()
End Sub
