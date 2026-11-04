Option Explicit
Private Declare PtrSafe Function Suite Lib "__DLL__" Alias "ForeignSuite" () As LongPtr
Private Declare PtrSafe Function Shutdown Lib "__DLL__" Alias "CalcVbaShutdown" () As Long
Private Declare PtrSafe Function BStrLength Lib "oleaut32" Alias "SysStringLen" (ByVal pointer As LongPtr) As Long
Private Declare PtrSafe Sub FreeBStr Lib "oleaut32" Alias "SysFreeString" (ByVal pointer As LongPtr)
Private Declare PtrSafe Sub CopyBytes Lib "kernel32" Alias "RtlMoveMemory" (ByVal destination As LongPtr, ByVal source As LongPtr, ByVal bytes As LongPtr)
Public Function ProbeNativeExports() As String
 Dim pointer As LongPtr, length As Long
 pointer=Suite(): length=BStrLength(pointer)
 ProbeNativeExports=String$(length,ChrW$(0))
 If length Then CopyBytes StrPtr(ProbeNativeExports),pointer,length*2&
 FreeBStr pointer
 If Shutdown()<>0& Then Err.Raise 513&
End Function
Public Sub Main()
 Dim result As String
 result=ProbeNativeExports()
End Sub
