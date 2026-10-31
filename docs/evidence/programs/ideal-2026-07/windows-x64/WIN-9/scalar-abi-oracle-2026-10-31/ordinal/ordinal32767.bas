Private Declare PtrSafe Function NativeMissing Lib "C:\Work\DnaCalc\codex-runs\calcvba-exact-native-probe\scalar.dll" Alias "#32767" () As Long
Private Declare PtrSafe Sub NativeSetError Lib "kernel32" Alias "SetLastError" (ByVal value As Long)
Public Function RunProbe() As String
    Dim value As Long, prior As Long
    NativeSetError 1234
    prior = Err.LastDllError
    On Error Resume Next
    value = NativeMissing()
    RunProbe = CStr(Err.Number) & "|" & Err.Description & "|" & Err.Source & "|" & CStr(prior) & "|" & CStr(Err.LastDllError)
End Function
