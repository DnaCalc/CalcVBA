Option Explicit
Private Declare PtrSafe Function Projection Lib "__FIXTURE__" Alias "FixtureProjection" (ByVal index As Long) As LongPtr
Private Declare PtrSafe Function References Lib "__FIXTURE__" Alias "FixtureReferences" () As Long
Private Declare PtrSafe Function SameObject Lib "__DLL__" (ByVal first As LongPtr, ByVal second As LongPtr) As Boolean
Private Declare PtrSafe Function Shutdown Lib "__DLL__" Alias "CalcVbaShutdown" () As Long
Public Function ProbeNativeExports() As String
    Dim first As LongPtr, alternate As LongPtr, second As LongPtr
    first = Projection(0&): alternate = Projection(1&): second = Projection(2&)
    If first = alternate Or References() <> 0& Then Err.Raise 513&
    ProbeNativeExports = CStr(CLng(SameObject(first, alternate))) & "|" & CStr(CLng(SameObject(first, second))) & "|" & CStr(CLng(SameObject(0^, 0^))) & "|" & CStr(References())
    If Shutdown() <> 0& Or References() <> 0& Then Err.Raise 514&
End Function
Public Sub Main()
    Dim result As String
    result = ProbeNativeExports()
End Sub
