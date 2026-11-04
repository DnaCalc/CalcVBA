VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Counter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mValue As Long
Public Property Get Value() As Long
    Value = mValue
End Property
Public Property Let Value(ByVal value As Long)
    mValue = value
End Property
