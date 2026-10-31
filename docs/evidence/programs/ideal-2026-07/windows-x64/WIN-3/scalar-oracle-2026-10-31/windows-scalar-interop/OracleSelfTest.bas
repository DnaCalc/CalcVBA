Option Explicit
Private Declare PtrSafe Function NativeSqrt Lib "msvcrt" Alias "sqrt" (ByVal value As Double) As Double
Private Declare PtrSafe Function VarR8FromI4 Lib "oleaut32" (ByVal value As Long, ByRef result As Double) As Long
Private Declare PtrSafe Sub NativeSetError Lib "kernel32" Alias "SetLastError" (ByVal value As Long)

Public Function Add2(a, b)
    Add2 = a + b
End Function

Public Function RunProbe() As String
    Dim result As Double
    Dim status As Long
    Dim savedError As Long
    Dim obj As Object
    status = VarR8FromI4(42, result)
    NativeSetError 1234
    savedError = Err.LastDllError
    Set obj = CreateObject("Scripting.Dictionary")
    obj.Add "answer", 42
    RunProbe = CStr(CLng(NativeSqrt(1764))) & "|" & CStr(result) & "|" & _
        CStr(savedError) & "|" & CStr(obj.Item("answer")) & "|" & _
        CStr(obj.Count) & "|" & CStr(Add2(19, 23))
End Function
