VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Nullable"
Attribute VB_Exposed = False
Option Explicit
Private child As Nullable
Private initialized As Boolean
Private Sub Class_Initialize()
    initialized=(child Is Nothing)
End Sub
Public Function Defaults() As String
    Defaults=CStr(VarType(child)) & ":" & CStr(child Is Nothing) & ":" & CStr(initialized)
End Function
Friend Function Internal(ByVal opcode As Long) As Long
    Internal=opcode+100
End Function
Private Function Hidden() As Long
    Hidden=1000
End Function
Public Function ReadInternal() As Long
    ReadInternal=Me.Internal(7)
End Function
