Attribute VB_Name = "ComExeClient"
Option Explicit
Public Function Probe() As String
    Dim first As CalcVbaNativeCounter27Lib.Counter, second As CalcVbaNativeCounter27Lib.Counter
    Dim value As Long, result As String, owner As Long
    Set first = New CalcVbaNativeCounter27Lib.Counter
    Set second = New CalcVbaNativeCounter27Lib.Counter
    result = CStr(first.Add(2, 3)) & ":" & CStr(second.Value)
    first.Value = 41
    value = 7
    result = result & ":" & CStr(first.Exchange(value)) & ":" & CStr(value) & ":" & first.Echo("hello")
    owner = first.ThreadId()
    result = result & ":" & CStr(owner = second.ThreadId())
    On Error Resume Next
    value = first.Fail()
    result = result & ":" & CStr(Err.Number)
    Dim late As Object
    Set late = first
    result = result & ":" & late.Echo("late")
    Probe = result
End Function
