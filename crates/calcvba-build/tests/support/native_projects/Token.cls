Attribute VB_Name = "Token"
Attribute VB_Exposed = True
Option Explicit
Private value As Long
Private Sub Class_Initialize()
    value = 3
End Sub
Private Sub Class_Terminate()
    Common.Terminated
End Sub
Public Sub PutValue(ByVal number As Long)
    value = number
End Sub
Public Function GetValue() As Long
    GetValue = value
End Function
