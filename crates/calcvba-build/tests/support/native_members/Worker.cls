VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Worker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICounter
Private caption As String
Private Function ICounter_Apply(ByRef value As Long) As Long
    value = value + 2
    ICounter_Apply = value * 3
End Function

Public Property Get Name() As String
    Name=caption
End Property
Public Property Let Name(ByVal value As String)
    caption=value
End Property
Public Function Echo() As String
    Echo="worker"
End Function
