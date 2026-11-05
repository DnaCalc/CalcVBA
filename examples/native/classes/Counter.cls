VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Counter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private count As Long
Private Sub Class_Initialize()
    Counts.created = Counts.created + 1
    count = 2
End Sub
Private Sub Class_Terminate()
    Counts.destroyed = Counts.destroyed + 1
End Sub
Public Property Get Value() As Long
    Value = count
End Property
Public Property Let Value(ByVal value As Long)
    count = value
End Property
Public Function Bump(ByRef delta As Long) As Long
    count = count + delta
    delta = delta + 1
    Bump = count
End Function
