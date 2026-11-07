VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Counter"
Attribute VB_Creatable = True
Attribute VB_Exposed = True
Option Explicit
Private Declare PtrSafe Function CurrentThread Lib "kernel32" Alias "GetCurrentThreadId" () As Long
Private total As Long
Private Sub Class_Initialize()
    total = 10
End Sub
Public Function Add(ByVal left As Long, ByVal right As Long) As Long
    total = total + left + right
    Add = total
End Function
Public Function Echo(ByVal text As String) As String
    Echo = text & ":" & CStr(total)
End Function
Public Function Exchange(ByRef value As Long) As Long
    Exchange = value
    value = total
    total = total + 1
End Function
Public Property Get Value() As Long
    Value = total
End Property
Public Property Let Value(ByVal value As Long)
    total = value
End Property
Public Function ThreadId() As Long
    ThreadId = CurrentThread()
End Function
Public Function Fail() As Long
    Err.Raise 513, "NativeCounter", "compiled failure"
End Function
