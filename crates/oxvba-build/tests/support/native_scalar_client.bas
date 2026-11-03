Option Explicit
Private Declare PtrSafe Function NativeAddLong Lib "__DLL__" Alias "AddLong" (ByVal a As Long, ByVal b As Long) As Long
Private Declare PtrSafe Function NativeDirectLong Lib "__DLL__" Alias "DirectLong" (ByVal a As Long, ByVal b As Long) As Long
Private Declare PtrSafe Function NativeAddDouble Lib "__DLL__" Alias "AddDouble" (ByVal a As Double, ByVal b As Double) As Double
Private Declare PtrSafe Function NativeGreater Lib "__DLL__" Alias "Greater" (ByVal a As Long, ByVal b As Long) As Boolean
Private Declare PtrSafe Function NativeEchoBoolean Lib "__DLL__" Alias "EchoBoolean" (ByVal value As Boolean) As Boolean
Private Declare PtrSafe Function NativeMixed5 Lib "__DLL__" Alias "Mixed5" (ByVal a As Double, ByVal b As Long, ByVal flag As Boolean, ByVal d As Double, ByVal e As Long) As Double
Private Declare PtrSafe Function NativeDivide Lib "__DLL__" Alias "Divide" (ByVal a As Double, ByVal b As Double) As Double
Private Declare PtrSafe Function NativeLibraryAbs Lib "__DLL__" Alias "LibraryAbs" (ByVal value As Double) As Double
Private Declare PtrSafe Function NativeLibrarySqr Lib "__DLL__" Alias "LibrarySqr" (ByVal value As Double) As Double

Public Function ProbeNativeExports() As String
    ProbeNativeExports = CStr(NativeAddLong(-7&, 49&)) & "|" & CStr(NativeDirectLong(5&, 32&))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeAddDouble(1.25, 2.5))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeGreater(5&, -2&))) & "|" & CStr(CLng(NativeGreater(-2&, 5&)))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(CLng(NativeEchoBoolean(True))) & "|" & CStr(CLng(NativeEchoBoolean(False)))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeMixed5(1.25, 2&, True, 3.5, 4&)) & "|" & CStr(NativeMixed5(1.25, 2&, False, 3.5, 4&))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeDivide(9#, 2#)) & "|" & CStr(NativeDivide(9#, 0#))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeAddLong(2147483647, 1&))
    ProbeNativeExports = ProbeNativeExports & "|" & CStr(NativeLibraryAbs(-12.5)) & "|" & CStr(NativeLibrarySqr(25#)) & "|" & CStr(NativeLibrarySqr(-1#))
End Function

Public Sub Main()
    Dim result As String
    result = ProbeNativeExports()
End Sub
